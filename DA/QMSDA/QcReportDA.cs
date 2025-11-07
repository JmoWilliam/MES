using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;
using static Dapper.SqlMapper;

namespace QMSDA
{
    public class QcReportDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string AutomationConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string HrmEtergeConnectionStrings = "";

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

        public QcReportDA()
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
        #region //GetDeliveryDetail -- 取得出貨單身紀錄 -- WuTc -- 2024-05-27
        public string GetDelveryDetail(string StartDate, string EndDate, int CustomerId, int DcId, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT b.DoDetailId, a.DoId, a.DoDate, g.MtlItemNo, g.MtlItemName, g.TypeThree
                        , a.[Status], j.StatusName DoStatus, FORMAT(a.DoDate, 'yyyy-MM-dd hh:mm') DoDateTime
	                    , ISNULL(h.CustomerMtlItemName, '') CustomerMtlItemName
	                    , i.SoErpPrefix + '-' + i.SoErpNo + '-' + c.SoSequence SoErpNo
	                    , ISNULL(b.DoQty, 0) NormalQty, ISNULL(b.FreebieQty, 0) FreebieQty, ISNULL(b.SpareQty, 0) SpareQty, b.OrderSituation, b.DeliveryMethod
                        , e.CustomerName, e.CustomerId DeliveryCustomerId
                        , d.DcShortName, d.DcId
                        FROM SCM.DeliveryOrder a
	                    INNER JOIN SCM.DoDetail b ON a.DoId = b.DoId
	                    INNER JOIN SCM.SoDetail c ON b.SoDetailId = c.SoDetailId
                        INNER JOIN SCM.DeliveryCustomer d ON a.DcId = d.DcId
                        INNER JOIN SCM.Customer e ON d.CustomerId = e.CustomerId
	                    INNER JOIN SCM.PickingItem f ON a.DoId = f.DoId AND  b.SoDetailId = f.SoDetailId
	                    INNER JOIN PDM.MtlItem g ON c.MtlItemId = g.MtlItemId
	                    LEFT JOIN PDM.CustomerMtlItem h ON g.MtlItemId = h.MtlItemId AND a.CustomerId = h.CustomerId
	                    INNER JOIN SCM.SaleOrder i ON c.SoId = i.SoId
                        INNER JOIN BAS.[Status] j ON a.[Status] = j.StatusNo AND a.[Status] != 'N' 
                                                    AND a.[Status] != 'P' AND j.StatusSchema = 'DeliveryOrder.Status'
                        WHERE 1=1";

                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DeliveryDate", @" AND FORMAT(a.DoDate, 'yyyy-MM-dd') = @DeliveryDate ", StartDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.DoDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.DoDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (MtlItemNo != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND g.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    if (MtlItemName != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND g.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    if (CustomerId != -1) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", @" AND d.CustomerId = @CustomerId", CustomerId);
                    if (DcId != -1) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DcId", @" AND d.DcId = @DcId", DcId);
                    sql += "    ORDER BY DoDateTime, SoErpNo, b.DoDetailId";

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

        #region //GetInventoryDetail -- 取得入庫單紀錄 -- WuTc -- 2024-07-23
        public string GetInventoryDetail(string StartDate, string EndDate, int CustomerId, string MoErpNo, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MweId, a.DocDate, a.Quantity, a.QcFlag, a.MweErpPrefix + '-' + a.MweErpNo MweErpNo
	                        , c.MtlItemId, c.MtlItemNo, c.MtlItemName, c.MtlItemDesc
	                        , ISNULL(Cus.CustomerId, '0') CustomerId, ISNULL(Cus.CustomerShortName, '製令未綁定單') CustomerShortName
	                        , d.MoId, e.WoErpPrefix + '-' + e.WoErpNo + '(' + CONVERT(VARCHAR, d.WoSeq) + ')' MoErpNo
	                        FROM MES.MoWarehouseEnry a
	                        INNER JOIN MES.MweDetail b ON a.MweId = b.MweId
	                        INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
	                        INNER JOIN MES.ManufactureOrder d ON b.MoId = d.MoId
	                        INNER JOIN MES.WipOrder e ON e.WoId = d.WoId
	                        OUTER APPLY ( SELECT mc.CustomerId, md.CustomerShortName FROM MES.WipOrder ma
		                        INNER JOIN SCM.SoDetail mb ON ma.SoDetailId = mb.SoDetailId
		                        INNER JOIN SCM.SaleOrder mc ON mb.SoId = mc.SoId
		                        INNER JOIN SCM.Customer md ON mc.CustomerId = md.CustomerId
		                        WHERE  d.WoId = ma.WoId ) Cus
                            WHERE a.CompanyId = @CompanyId";

                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (MtlItemNo != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND c.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    if (MtlItemName != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND c.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    if (MoErpNo != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoErpNo", @" AND e.WoErpPrefix + '-' + e.WoErpNo + '(' + CONVERT(VARCHAR, d.WoSeq) + ')' = @MoErpNo", MoErpNo);
                    if (CustomerId != -1) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", @" AND Cus.CustomerId = @CustomerId", CustomerId);
                    sql += "    ORDER BY MoErpNo, MtlItemNo, MweErpNo";

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

        #region //GetCustomerReport -- 取得客戶報表格式 -- WuTc 2024-05-28
        public string GetCustomerReport(int CustomerId, int ProductType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DcrId, ProductType, CustomerId, FunctionName, EmptyReportName, Encoding
                            FROM SCM.DeliveryCustomerReport 
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", @" AND CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProductType", @" AND ProductType = @ProductType", ProductType);

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

        #region //GetDeliveryMeasurementData -- 取得出貨量測紀錄 -- WuTc -- 2024-05-27
        public string GetDeliveryMeasurementData(string MweId, string InventoryDate, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //檢查是否有允收標準
                    //若有設定允收標準，則項目及規格依照允收標準的來
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT f.QcItemNo, f.QcItemName, f.QcItemDesc, f.QcItemId, e.QcItemDesc
	                        FROM PDM.MtlItem a
	                        INNER JOIN MES.WipOrder b ON a.MtlItemId = b.MtlItemId
	                        INNER JOIN MES.ManufactureOrder c ON b.WoId = c.WoId
	                        INNER JOIN MES.MweDetail d ON c.MoId = d.MoId
	                        INNER JOIN PDM.MtlQcItem e ON b.MtlItemId = e.MtlItemId
	                        INNER JOIN QMS.QcItem f ON e.QcItemId = f.QcItemId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MweId", @" AND d.MweId = @MweId", MweId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND a.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    sql += "    ORDER BY f.QcItemNo";

                    var mtlQcitemresult = sqlConnection.Query(sql, dynamicParameters);

                    string QcItemNo = "";
                    if (mtlQcitemresult.Count() > 0) //允收標準量測項目
                    {
                        foreach (var item in mtlQcitemresult)
                        {
                            if (QcItemNo != "") QcItemNo += ',';
                            string qcItemNo = item.QcItemNo.ToString();
                            QcItemNo += "'" + qcItemNo + "'";
                        }
                    }
                    #endregion

                    #region //取得出貨檢量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT qc.QmdId, a.MweId, a.DocDate, a.Quantity, a.MweErpPrefix + '-' + a.MweErpNo MweErpNo
	                            , Cus.CustomerId, Cus.CustomerShortName, Cus.CustomerMtlItemNo
	                            , mt.MtlItemId, mt.MtlItemNo, mt.MtlItemName, mt.MtlItemDesc
	                            , ( SELECT ccd.CustomerDwgNo
		                            FROM PDM.CustomerCad ccd
		                            WHERE ccd.CustomerMtlItemNo = Cus.CustomerMtlItemNo ) CustomerDwgNo
	                            , qc.WoId, qc.MoId, qc.WoSeq, qc.MakeCount
	                            , qc.BarcodeId AS QcBarcodeId, bc.ItemValue, bc.ItemSeq
	                            , qc.Cavity, qc.LotNumber, qc.QcRecordId
	                            , qc.QcItemId, qc.QcItemDesc, qc.QmmDetailId
	                            , qc.QcItemNo, qc.QcClassNo, qc.QcClassName
	                            , ISNULL(CONVERT(nvarchar(18), qc.DesignValue), '') AS DesignValue
	                            , ISNULL(CONVERT(nvarchar(18), qc.UpperTolerance), '') AS UpperTolerance
	                            , ISNULL(CONVERT(nvarchar(18), qc.LowerTolerance), '') AS LowerTolerance
	                            , ISNULL(qc.ZAxis, '') ZAxis, qc.CauseId, qc.MeasureValue
	                            , qc.BallMark, qc.Unit
	                            , qc.InputType, qc.QcTypeName
	                            , MM.MachineId, MM.MachineNo, MM.MachineNumber, MM.MachineName, MM.MachineDesc
	                            , qc.AbnormalqualityNo, qc.AbnormalqualityStatus, qc.QcType
	                            , qc.MtlQcItemId, qc.MtlQcItemDesc, qc.MtlDesignValue, qc.MtlUpperTolerance, qc.MtlLowerTolerance, qc.MtlBallMark, qc.MtlUnit
	                            , (SELECT bta.TypeName FROM BAS.[Type] bta
		                            WHERE bta.TypeNo = bc.CurrentProdStatus
		                            AND bta.TypeSchema = 'Barcode.CurrentProdStatus' ) CurrentProdStatus
	                            , ( SELECT btb.TypeName FROM BAS.[Type] btb
		                            WHERE btb.TypeNo = bc.BarcodeStatus
		                            AND btb.TypeSchema = 'Barcode.BarcodeStatus' ) BarcodeStatus
	                            , ( SELECT af.UserName FROM BAS.[User] af
		                            WHERE af.UserId = qc.QcUserId ) QcUserName
	                            , ln.LotNumberNo, ln.CloseStatus
	                            FROM MES.MoWarehouseEnry a
	                            INNER JOIN MES.MweDetail md ON md.MweId = a.MweId
	                            INNER JOIN MES.MweBarcode mb ON mb.MweDetailId = md.MweDetailId
	                            INNER JOIN PDM.MtlItem mt ON mt.MtlItemId = md.MtlItemId
	                            LEFT JOIN SCM.LotNumber ln ON mt.MtlItemId = ln.MtlItemId
	                            OUTER APPLY (
		                            SELECT d.WoId, d.MoId, d.WoSeq, g.QcRecordId, g.InputType, f.BarcodeId, qt.QcTypeName
		                            , f.QcItemId, f.Cavity, f.LotNumber, f.QmdId, f.QcItemDesc, f.QmmDetailId, f.BallMark, f.Unit
		                            , f.DesignValue, f.UpperTolerance, f.LowerTolerance, f.ZAxis, f.CauseId, f.MeasureValue
		                            , g.QcUserId, wo.SoDetailId, f.MakeCount
		                            , qi.QcItemNo, qcl.QcClassNo, qcl.QcClassName
		                            , aq.AbnormalqualityNo, aq.AbnormalqualityStatus, aq.QcType
		                            , g.CreateDate, g.LastModifiedDate
                                    , mq.QcItemId AS MtlQcItemId, mq.QcItemDesc AS MtlQcItemDesc, mq.DesignValue AS MtlDesignValue, mq.UpperTolerance AS MtlUpperTolerance, mq.LowerTolerance AS MtlLowerTolerance
                                    , mq.BallMark AS MtlBallMark, mq.Unit AS MtlUnit
		                            FROM MES.ManufactureOrder d
		                            INNER JOIN MES.WipOrder wo ON d.WoId = wo.WoId
		                            INNER JOIN MES.QcRecord g ON d.MoId = g.MoId
		                            INNER JOIN QMS.QcMeasureData f ON f.QcRecordId = g.QcRecordId
		                            INNER JOIN QMS.QcType qt ON g.QcTypeId = qt.QcTypeId AND NOT (qt.QcTypeName LIKE '%工程%')
		                            INNER JOIN QMS.QcItem qi ON f.QcItemId = qi.QcItemId
		                            INNER JOIN QMS.QcClass qcl ON qcl.QcClassId = qi.QcClassId
		                            LEFT JOIN QMS.Abnormalquality aq ON g.MoId = aq.MoId AND aq.QcType != 'NON'	
		                            LEFT JOIN PDM.MtlQcItem mq ON qi.QcItemId = mq.QcItemId
		                            WHERE d.MoId = md.MoId AND f.BarcodeId = mb.BarcodeId
		                            ) qc
	                            OUTER APPLY (
		                            SELECT DISTINCT bb.BarcodeId, bb.BarcodeNo, bb.BarcodeStatus
		                            , bb.CurrentProdStatus, ba.ItemValue ,ba.ItemSeq
		                            FROM MES.Barcode bb 
		                            LEFT JOIN MES.BarcodeProcess bp ON bb.BarcodeId = bp.BarcodeId AND bp.ProdStatus = 'P'
		                            LEFT JOIN MES.BarcodeAttribute ba ON bb.BarcodeId = ba.BarcodeId
		                            WHERE mb.BarcodeId = bb.BarcodeId AND bb.BarcodeStatus != '1'
		                            ) bc
	                            OUTER APPLY ( 
		                            SELECT so.CustomerId, cs.CustomerShortName, sd.CustomerMtlItemNo
		                            FROM SCM.SoDetail sd
		                            INNER JOIN SCM.SaleOrder so ON so.SoId = sd.SoId
		                            INNER JOIN SCM.Customer cs ON cs.CustomerId = so.CustomerId
		                            WHERE  sd.SoDetailId = qc.SoDetailId
		                            ) Cus
	                            OUTER APPLY	( 
		                            SELECT qd.MachineNumber, me.MachineNo, me.MachineName, me.MachineDesc, me.MachineId
		                            FROM MES.Machine me
		                            INNER JOIN QMS.QmmDetail qd ON qd.MachineId = me.MachineId
		                            WHERE qd.QmmDetailId = qc.QmmDetailId
		                            ) MM
                            WHERE qc.QmdId != '' AND (FORMAT(qc.LastModifiedDate,'yyyy-mm-dd hh:ss') >= (SELECT FORMAT(MAX(LastModifiedDate),'yyyy-mm-dd hh:ss') FROM QMS.QcMeasureData WHERE QcRecordId = qc.QcRecordId))";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "InventoryDate", @" AND a.DocDate = @InventoryDate ", InventoryDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND mt.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND mt.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MweId", @" AND a.MweId = @MweId", MweId);
                    if (QcItemNo.Length > 0)
                    {
                        sql += @" AND qc.QcItemNo IN (" + QcItemNo + ")";
                    }

                    sql += @" UNION ALL
                                SELECT DISTINCT qc.QmdId, a.MweId, a.DocDate, a.Quantity, a.MweErpPrefix + '-' + a.MweErpNo MweErpNo
	                                , Cus.CustomerId, Cus.CustomerShortName, Cus.CustomerMtlItemNo
	                                , mt.MtlItemId, mt.MtlItemNo, mt.MtlItemName, mt.MtlItemDesc
	                                , ( SELECT ccd.CustomerDwgNo
		                                FROM PDM.CustomerCad ccd
		                                WHERE ccd.CustomerMtlItemNo = Cus.CustomerMtlItemNo ) CustomerDwgNo
	                                , qc.WoId, qc.MoId, qc.WoSeq, qc.MakeCount
	                                , qc.BarcodeId AS QcBarcodeId, bc.ItemValue, bc.ItemSeq
	                                , qc.Cavity, qc.LotNumber, qc.QcRecordId
	                                , qc.QcItemId, qc.QcItemDesc, qc.QmmDetailId
	                                , qc.QcItemNo, qc.QcClassNo, qc.QcClassName
	                                , ISNULL(CONVERT(nvarchar(18), qc.DesignValue), '') AS DesignValue
	                                , ISNULL(CONVERT(nvarchar(18), qc.UpperTolerance), '') AS UpperTolerance
	                                , ISNULL(CONVERT(nvarchar(18), qc.LowerTolerance), '') AS LowerTolerance
	                                , ISNULL(qc.ZAxis, '') ZAxis, qc.CauseId, qc.MeasureValue
	                                , qc.BallMark, qc.Unit
	                                , qc.InputType, qc.QcTypeName
	                                , MM.MachineId, MM.MachineNo, MM.MachineNumber, MM.MachineName, MM.MachineDesc
	                                , qc.AbnormalqualityNo, qc.AbnormalqualityStatus, qc.QcType
	                                , qc.MtlQcItemId, qc.MtlQcItemDesc, qc.MtlDesignValue, qc.MtlUpperTolerance, qc.MtlLowerTolerance, qc.MtlBallMark, qc.MtlUnit
	                                , (SELECT bta.TypeName FROM BAS.[Type] bta
		                                WHERE bta.TypeNo = bd.CurrentProdStatus
		                                AND bta.TypeSchema = 'Barcode.CurrentProdStatus' ) CurrentProdStatus
	                                , ( SELECT btb.TypeName FROM BAS.[Type] btb
		                                WHERE btb.TypeNo = bd.BarcodeStatus
		                                AND btb.TypeSchema = 'Barcode.BarcodeStatus' ) BarcodeStatus
	                                , ( SELECT af.UserName FROM BAS.[User] af
		                                WHERE af.UserId = qc.QcUserId ) QcUserName
	                                , ln.LotNumberNo, ln.CloseStatus
	                                FROM MES.MoWarehouseEnry a
	                                INNER JOIN MES.MweDetail md ON md.MweId = a.MweId
	                                LEFT JOIN MES.MweBarcode mb ON mb.MweDetailId = md.MweDetailId
	                                INNER JOIN PDM.MtlItem mt ON mt.MtlItemId = md.MtlItemId
	                                LEFT JOIN SCM.LotNumber ln ON mt.MtlItemId = ln.MtlItemId
	                                OUTER APPLY (
		                                SELECT d.WoId, d.MoId, d.WoSeq, g.QcRecordId, g.InputType, f.BarcodeId, qt.QcTypeName
		                                , f.QcItemId, f.Cavity, f.LotNumber, f.QmdId, f.QcItemDesc, f.QmmDetailId, f.BallMark, f.Unit
		                                , f.DesignValue, f.UpperTolerance, f.LowerTolerance, f.ZAxis, f.CauseId, f.MeasureValue
		                                , g.QcUserId, wo.SoDetailId, f.MakeCount
		                                , qi.QcItemNo, qcl.QcClassNo, qcl.QcClassName
		                                , aq.AbnormalqualityNo, aq.AbnormalqualityStatus, aq.QcType
		                                , g.CreateDate, g.LastModifiedDate
                                        , mq.QcItemId AS MtlQcItemId, mq.QcItemDesc AS MtlQcItemDesc, mq.DesignValue AS MtlDesignValue, mq.UpperTolerance AS MtlUpperTolerance, mq.LowerTolerance AS MtlLowerTolerance
                                        , mq.BallMark AS MtlBallMark, mq.Unit AS MtlUnit
		                                FROM MES.ManufactureOrder d
		                                INNER JOIN MES.WipOrder wo ON d.WoId = wo.WoId
		                                INNER JOIN MES.QcRecord g ON d.MoId = g.MoId
		                                INNER JOIN QMS.QcMeasureData f ON f.QcRecordId = g.QcRecordId
		                                INNER JOIN QMS.QcType qt ON g.QcTypeId = qt.QcTypeId --AND NOT (qt.QcTypeName LIKE '%工程%')
		                                INNER JOIN QMS.QcItem qi ON f.QcItemId = qi.QcItemId
		                                INNER JOIN QMS.QcClass qcl ON qcl.QcClassId = qi.QcClassId
		                                LEFT JOIN QMS.Abnormalquality aq ON g.MoId = aq.MoId AND aq.QcType != 'NON'	
		                                LEFT JOIN PDM.MtlQcItem mq ON qi.QcItemId = mq.QcItemId
		                                WHERE d.MoId = md.MoId AND f.LotNumber = ln.LotNumberNo
		                                ) qc
	                                LEFT JOIN MES.Barcode bd ON mb.BarcodeId = bd.BarcodeId
	                                LEFT JOIN MES.BarcodeAttribute bc ON bd.BarcodeId = bc.BarcodeId
	                                OUTER APPLY ( 
		                                SELECT so.CustomerId, cs.CustomerShortName, sd.CustomerMtlItemNo
		                                FROM SCM.SoDetail sd
		                                INNER JOIN SCM.SaleOrder so ON so.SoId = sd.SoId
		                                INNER JOIN SCM.Customer cs ON cs.CustomerId = so.CustomerId
		                                WHERE  sd.SoDetailId = qc.SoDetailId
		                                ) Cus
	                                OUTER APPLY	( 
		                                SELECT qd.MachineNumber, me.MachineNo, me.MachineName, me.MachineDesc, me.MachineId
		                                FROM MES.Machine me
		                                INNER JOIN QMS.QmmDetail qd ON qd.MachineId = me.MachineId
		                                WHERE qd.QmmDetailId = qc.QmmDetailId
		                                ) MM
	                            WHERE qc.QmdId != '' AND (FORMAT(qc.LastModifiedDate,'yyyy-mm-dd hh:ss') >= (SELECT FORMAT(MAX(LastModifiedDate),'yyyy-mm-dd hh:ss') FROM QMS.QcMeasureData WHERE QcRecordId = qc.QcRecordId))";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "InventoryDate", @" AND a.DocDate = @InventoryDate ", InventoryDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND mt.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND mt.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MweId", @" AND a.MweId = @MweId", MweId);
                    if (QcItemNo.Length > 0)
                    {
                        sql += @" AND qc.QcItemNo IN (" + QcItemNo + ")";
                    }

                    sql += "    ORDER BY qc.QcItemNo, qc.QmmDetailId, bc.ItemSeq, qc.Cavity, qc.LotNumber, qc.BarcodeId";

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

        #region //GetMeasurementDataFile -- 取得量測紀錄上傳歸檔的檔案 -- WuTc -- 2024-10-11
        public string GetMeasurementDataFile(string QcRecordId, string inputType, List<string> pcsNumberList)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得量測數據上傳歸檔的檔案
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcRecordId, a.FileId, a.QcRecordFileId, a.FileType, a.PhysicalPath, a.InputType, a.BarcodeId, a.Cavity, a.LotNumber, a.EffectiveDiameter, c.ItemSeq
                            FROM MES.QcRecordFile a
                            INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                            LEFT JOIN MES.BarcodeAttribute c ON a.BarcodeId = c.BarcodeId
                            WHERE b.QcRecordId = @QcRecordId AND a.FileType != 'upload-ModityExcel'
                            ";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    if (pcsNumberList.Count > 0)
                    {
                        if (inputType == "Lettering")
                        {
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItemSeq", @" AND c.ItemSeq IN @ItemSeq", pcsNumberList); //刻號
                        }
                        else if (inputType == "Cavity")
                        {
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Cavity", @" AND a.Cavity IN @Cavity", pcsNumberList); //穴號
                        }
                        else if (inputType == "LotNumber")
                        {
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "LotNumber", @" AND a.LotNumber IN @LotNumber", pcsNumberList); //批號
                        }
                        else if (inputType == "BarcodeId")
                        {
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BarcodeId", @" AND a.BarcodeId IN @BarcodeId", pcsNumberList); //條碼
                        }
                    }

                    sql += "    ORDER BY a.PhysicalPath, a.BarcodeId, a.Cavity, a.LotNumber, a.EffectiveDiameter";
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

        #region //GetModifyMeasureData -- 取得修改後上傳的excel筆數 -- WuTc -- 2024-08-09
        public string GetModifyMeasureData(string StartDate, string EndDate, string MoIdErpNo, string MtlItemNo, string MtlItemName, string UserNo, int CustomerId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT COUNT(DISTINCT LetteringSeq) Quantity, a.MoId, FORMAT(a.CreateDate,'yyyy-MM-dd') CreateDate
                            , d.WoErpPrefix + '-' + d.WoErpNo + '(' + CONVERT(VARCHAR, c.WoSeq) + ')' MoErpNo
	                        , f.UserNo, f.UserName
	                        , e.MtlItemNo, e.MtlItemName
                            , Cus.CustomerId, Cus.CustomerShortName
	                        FROM QMS.QcMeasureDataModify a
	                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
	                        INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
	                        INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
	                        INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
	                        INNER JOIN BAS.[User] f ON a.CreateBy = f.UserId
                            OUTER APPLY ( SELECT mc.CustomerId, md.CustomerShortName FROM MES.WipOrder ma
		                        INNER JOIN SCM.SoDetail mb ON ma.SoDetailId = mb.SoDetailId
		                        INNER JOIN SCM.SaleOrder mc ON mb.SoId = mc.SoId
		                        INNER JOIN SCM.Customer md ON mc.CustomerId = md.CustomerId
		                        WHERE  d.WoId = ma.WoId ) Cus
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND e.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserNo", @" AND f.UserNo = @UserNo", UserNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", @" AND Cus.CustomerId = @CustomerId", CustomerId);

                    if (MoIdErpNo != "")
                    {
                        if (MoIdErpNo.IndexOf('-') != -1)
                        {
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoIdErpNo", @" AND d.WoErpPrefix + '-' + d.WoErpNo + '(' + CONVERT(VARCHAR, c.WoSeq) + ')' = @MoIdErpNo", MoIdErpNo);
                        }
                        else
                        {
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoIdErpNo", @" AND a.MoId = @MoIdErpNo", Convert.ToInt32(MoIdErpNo));
                        }
                    }

                    sql += "    GROUP BY a.MoId, a.CreateDate, d.WoErpPrefix, d.WoErpNo, c.WoSeq, e.MtlItemNo, e.MtlItemName, f.UserNo, f.UserName, Cus.CustomerId, Cus.CustomerShortName" +
                            "   ORDER BY a.CreateDate DESC";

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

        #region //GetFileInfo -- 取得檔案資料 -- WuTc -- 2024-08-12
        public string GetFileInfo(string FileIdList, string FileType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FileId, a.[FileName], a.FileExtension
                            , a.[FileName] + a.FileExtension FileFullName, a.FileSize, a.FileContent
                            FROM BAS.[File] a
                            WHERE a.FileId IN (" + FileIdList + ")";

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("查無上傳的檔案資料！");

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

        #region //GetDeliveryMeasurementDataModify -- 取得修改後的出貨量測紀錄 -- WuTc -- 2024-08-16
        public string GetDeliveryMeasurementDataModify(string MweId, string UploadDate, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得出貨檢量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcRecordId, a.QmdId, a.CreateDate, b.MoId, d.WoId, c.WoSeq, a.Confirmer, a.Surveyor
	                        , a.LetteringSeq, a.Cavity, a.LotNumber, a.BarcodeId
	                        , Cus.CustomerId, Cus.CustomerShortName, Cus.CustomerMtlItemNo
	                        , e.MtlItemId, e.MtlItemNo, e.MtlItemName, e.MtlItemDesc
	                        , ( SELECT ccd.CustomerDwgNo 
		                        FROM PDM.CustomerCad ccd 
		                        WHERE ccd.CustomerMtlItemNo = Cus.CustomerMtlItemNo ) CustomerDwgNo		
	                        , barcode.BarcodeId, barcode.BarcodeNo, barcode.BarcodeStatus, barcode.ItemValue, barcode.ItemSeq
	                        , a.Cavity, a.LotNumber
	                        , a.QcItemId, a.QcItemDesc, a.QmmDetailId
	                        , QcItem.QcItemNo, QcItem.QcClassNo, QcItem.QcClassName
	                        , ISNULL(CONVERT(nvarchar(18), a.DesignValue), '') AS DesignValue
	                        , ISNULL(CONVERT(nvarchar(18), a.UpperTolerance), '') AS UpperTolerance
	                        , ISNULL(CONVERT(nvarchar(18), a.LowerTolerance), '') AS LowerTolerance
	                        , ISNULL(a.ZAxis, '') ZAxis, a.ModifyValue, a.BallMark, a.Unit
	                        , b.InputType
	                        , MM.MachineId, MM.MachineNo, MM.MachineNumber, MM.MachineName, MM.MachineDesc
	                        , f.UserNo, f.UserName
	                        FROM QMS.QcMeasureDataModify a
	                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
	                        INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
	                        INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
	                        INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
	                        INNER JOIN BAS.[User] f ON a.CreateBy = f.UserId
	                        OUTER APPLY ( SELECT DISTINCT bd.BarcodeId, bd.BarcodeNo, bd.BarcodeStatus
		                        , bd.CurrentProdStatus, ba.ItemValue ,ba.ItemSeq
		                        FROM MES.BarcodeProcess bp
		                        INNER JOIN MES.Barcode bd ON bd.BarcodeId = bp.BarcodeId
		                        INNER JOIN MES.BarcodeAttribute ba ON ba.BarcodeId = bd.BarcodeId
		                        WHERE bp.BarcodeId = a.BarcodeId AND bp.ProdStatus = 'P' AND bd.BarcodeStatus != '1') barcode
	                        OUTER APPLY ( SELECT so.CustomerId, cs.CustomerShortName, sd.CustomerMtlItemNo
		                        FROM MES.WipOrder wo
		                        INNER JOIN SCM.SoDetail sd ON sd.SoDetailId = wo.SoDetailId
		                        INNER JOIN SCM.SaleOrder so ON so.SoId = sd.SoId
		                        INNER JOIN SCM.Customer cs ON cs.CustomerId = so.CustomerId
		                        WHERE  wo.WoId = d.WoId ) Cus
	                        OUTER APPLY ( SELECT qi.QcItemNo, qc.QcClassNo, qc.QcClassName
		                        FROM QMS.QcItem qi 
		                        INNER JOIN QMS.QcClass qc ON qc.QcClassId = qi.QcClassId
		                        WHERE qi.QcItemId = a.QcItemId ) QcItem
	                        OUTER APPLY	( SELECT qd.MachineNumber, me.MachineNo, me.MachineName, me.MachineDesc, me.MachineId
		                        FROM MES.Machine me 
		                        INNER JOIN QMS.QmmDetail qd ON qd.MachineId = me.MachineId
		                        WHERE qd.QmmDetailId = a.QmmDetailId ) MM
	                        WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UploadDate", @" AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @UploadDate ", UploadDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND e.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MweId", @" AND a.MweId = @MweId", MweId);
                    sql += "    ORDER BY QcItem.QcItemNo, barcode.ItemSeq, a.Cavity, a.LotNumber, barcode.BarcodeId";
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("找不到量測資料！");
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

        #region //GetGoodsReceiptReportDataList -- 取得進貨檢驗量測紀錄列表 -- WuTc -- 2024-07-23
        public string GetGoodsReceiptReportDataList(string StartDate, string EndDate, int SupplierId, string GrFullErpNo, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcRecordId, a.InputType, a.QcStatus, FORMAT( a.CreateDate, 'yyyy-MM-dd  hh:mm') MeasureDate, q.QcTypeNo
	                        , c.MtlItemNo, c.GrMtlItemName
	                        , c.GrErpPrefix, c.GrErpNo, c.GrSeq, c.GrErpPrefix + '-' + c.GrErpNo + '(' + c.GrSeq + ')' GrFullErpNo
	                        , c.ReceiptQty, c.QcType
	                        , d.SupplierId, d.SupplierNo, d.SupplierName, s.SupplierShortName, FORMAT( d.ReceiptDate, 'yyyy-MM-dd ') ReceiptDate
	                        FROM MES.QcRecord a
	                        INNER JOIN MES.QcGoodsReceipt b ON a.QcRecordId = b.QcRecordId
	                        INNER JOIN SCM.GrDetail c ON b.GrDetailId = c.GrDetailId
	                        INNER JOIN SCM.GoodsReceipt d ON c.GrId = d.GrId
	                        INNER JOIN QMS.QcType q ON a.QcTypeId = q.QcTypeId AND q.QcTypeNo = 'IQC'
	                        INNER JOIN SCM.Supplier s ON d.SupplierId = s.SupplierId
                            WHERE d.CompanyId = @CompanyId";

                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND d.ReceiptDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND d.ReceiptDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (MtlItemNo != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND c.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    if (MtlItemName != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND c.GrMtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    if (GrFullErpNo != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrFullErpNo", @" AND c.GrErpPrefix + '-' + c.GrErpNo + '(' + c.GrSeq + ')' = @GrFullErpNo", GrFullErpNo);
                    if (SupplierId != -1) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SupplierId", @" AND d.SupplierId = @SupplierId", SupplierId);
                    sql += "    ORDER BY GrFullErpNo, MtlItemNo, c.AcceptanceDate";

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

        #region //GetGoodsReceiptData -- 取得進貨檢驗量測紀錄 -- WuTc -- 2024-10-05
        public string GetGoodsReceiptData(string QcRecordId, string GrDetailId, string MeasureDate, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //檢查是否有允收標準
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT e.QcItemId, e.QcItemNo, e.QcItemName, e.QcItemDesc, e.QcItemDesc
	                        FROM PDM.MtlItem a
	                        INNER JOIN SCM.GrDetail b ON a.MtlItemId = b.MtlItemId
	                        INNER JOIN MES.QcGoodsReceipt c ON b.GrDetailId = c.GrDetailId
	                        INNER JOIN PDM.MtlQcItem d ON a.MtlItemId = d.MtlItemId
	                        INNER JOIN QMS.QcItem e ON d.QcItemId = e.QcItemId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrDetailId", @" AND c.GrDetailId = @GrDetailId", GrDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND a.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    sql += "    ORDER BY f.QcItemNo";

                    var mtlQcitemresult = sqlConnection.Query(sql, dynamicParameters);

                    List<string> QcItemNo = new List<string>();
                    if (mtlQcitemresult.Count() > 0) //允收標準量測項目
                    {
                        foreach (var item in mtlQcitemresult)
                        {
                            QcItemNo.Add(item.QcItemNo);
                        }
                    }
                    #endregion

                    #region //取得進貨檢驗量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT c.GrDetailId, a.QcRecordId, a.QmdId, a.CreateDate MeasureDate
	                        , a.Cavity, a.LotNumber, a.BarcodeId
	                        , a.QcItemId, a.QcItemDesc
	                        , QcItem.QcItemNo, QcItem.QcClassNo, QcItem.QcClassName
	                        , ISNULL(CONVERT(nvarchar(18), a.DesignValue), '') AS DesignValue
	                        , ISNULL(CONVERT(nvarchar(18), a.UpperTolerance), '') AS UpperTolerance
	                        , ISNULL(CONVERT(nvarchar(18), a.LowerTolerance), '') AS LowerTolerance
	                        , ISNULL(a.ZAxis, '') ZAxis, a.MeasureValue, a.BallMark, a.Unit	
	                        , b.InputType
	                        , a.QmmDetailId, MM.MachineId, MM.MachineNo, MM.MachineNumber, MM.MachineName, MM.MachineDesc
	                        , d.MtlItemNo, d.GrMtlItemName
	                        , d.GrErpPrefix, d.GrErpNo, d.GrSeq, d.GrErpPrefix + '-' + d.GrErpNo + '(' + d.GrSeq + ')' GrFullErpNo
	                        , d.QcType, d.ReceiptQty, d.AcceptanceDate
	                        , e.SupplierId, e.SupplierNo, e.SupplierName
	                        , f.UserName, f.UserNo
	                        FROM QMS.QcMeasureData a
	                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
	                        INNER JOIN MES.QcGoodsReceipt c ON a.QcRecordId = c.QcRecordId
	                        INNER JOIN SCM.GrDetail d ON c.GrDetailId = d.GrDetailId
	                        INNER JOIN SCM.GoodsReceipt e ON d.GrId = e.GrId
	                        INNER JOIN BAS.[User] f ON b.QcUserId = f.UserId
	                        OUTER APPLY ( SELECT qi.QcItemNo, qc.QcClassNo, qc.QcClassName
		                        FROM QMS.QcItem qi 
		                        INNER JOIN QMS.QcClass qc ON qc.QcClassId = qi.QcClassId
		                        WHERE qi.QcItemId = a.QcItemId ) QcItem
	                        OUTER APPLY	( SELECT qd.MachineNumber, me.MachineNo, me.MachineName, me.MachineDesc, me.MachineId
		                        FROM MES.Machine me 
		                        INNER JOIN QMS.QmmDetail qd ON qd.MachineId = me.MachineId
		                        WHERE qd.QmmDetailId = a.QmmDetailId ) MM
	                        WHERE a.QcRecordId = @QcRecordId";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CreateDate", @" AND a.CreateDate = @CreateDate ", MeasureDate);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND c.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND c.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrDetailId", @" AND c.GrDetailId = @GrDetailId", GrDetailId);
                    if (QcItemNo.Count() > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcitemNo", @" AND a.QcitemNo IN @QcItemNo", QcItemNo);
                    sql += "    ORDER BY QcItem.QcItemNo, f.LotNumber";
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

        #region //GetSupplier -- 取得供應商資料 -- WuTc -- 2024-10-05
        public string GetSupplier(string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT SupplierId, SupplierNo, SupplierName, SupplierShortName 
                            FROM SCM.Supplier
                            WHERE [Status] = @Status";

                    dynamicParameters.Add("Status", Status);
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

        #region //GetMeasurementPointData -- 取得分光點資料量測紀錄 -- WuTc -- 2024-12-17
        public string GetMeasurementPointData(string mtlItemId, string moId, string QcRecordId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //檢查品號是否有設定分光的允收標準
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.QcItemNo, c.MtlItemNo, c.MtlItemName
                            , a.DesignValue, a.UpperTolerance, a.LowerTolerance, a.QcItemDesc, a.BallMark, a.Unit, a.Regulation, a.Graphics, a.Notice
                            , Qmm.MachineNumber, Qmm.MachineName
                            FROM PDM.MtlQcItem a
                            INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                            INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                            OUTER APPLY (
	                            SELECT qd.MachineNumber, mc.MachineName FROM QMS.QmmDetail qd
	                            INNER JOIN MES.Machine mc ON qd.MachineId = mc.MachineId
	                            WHERE qd.QmmDetailId = a.QmmDetailId
                            ) Qmm
                            WHERE b.QcItemNo LIKE '%s%' ";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId ", mtlItemId);
                    sql += "    ORDER BY b.QcItemNo";

                    var mtlQcitemresult = sqlConnection.Query(sql, dynamicParameters);
                    #endregion

                    List<string> QcItemNo = new List<string>();
                    if (mtlQcitemresult.Count() > 0) //允收標準量測項目
                    {
                        #region //判斷是否有分光點資料的數據路徑
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordFileId, a.PhysicalPath, a.FileType, ISNULL(a.BarcodeId, 0) BarcodeId, a.Cavity, a.EffectiveDiameter, a.LotNumber
                                FROM MES.QcRecordFile a
                                INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                WHERE a.FileType != 'upload-ModityExcel' AND (FORMAT(b.LastModifiedDate,'yyyy-mm-dd hh:ss') >= (SELECT FORMAT(MAX(LastModifiedDate),'yyyy-mm-dd hh:ss') FROM QMS.QcMeasureData WHERE QcRecordId = b.QcRecordId)) ";

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoId", @" AND b.MoId = @MoId ", moId);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcRecordId", @" AND b.QcRecordId = @QcRecordId ", QcRecordId);
                        sql += "    ORDER BY a.QcRecordFileId";

                        var pathresult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        if (pathresult.Count() <= 0) { throw new Exception("找不到點資料數據！"); }

                        List<string> result = new List<string>();
                        List<string> judge = new List<string>();
                        foreach (var item in pathresult)
                        {
                            int QcRecordFileId = item.QcRecordFileId;

                            #region //判斷是否有分光點資料的解析數據
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordFileId, a.PhysicalPath, a.FileType, ISNULL(a.BarcodeId, 0) BarcodeId, a.Cavity, a.EffectiveDiameter, a.LotNumber
                                , b.Point, b.PointSite, b.PointValue
                                FROM MES.QcRecordFile a
                                INNER JOIN QMS.QcMeasurePointData b ON a.QcRecordFileId = b.QcRecordFileId
                                WHERE a.FileType != 'upload-ModityExcel' AND a.QcRecordFileId = @QcRecordFileId ";
                            dynamicParameters.Add("QcRecordFileId", QcRecordFileId);

                            var pointresult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //依允收標準的規格，來取得點資料的值
                            //DesignValue 放的是點的位置 Point Wave，EX:420、380~460
                            //UpperTolerance 放的是取值種類及範圍 PointValue，分為正/反 PointSite，EX:Rabs.<、Rmax.<=、Tabs.>=
                            //LowerTolerance 放的是規格 0.5%、98.88%
                            //R是反射率，T是穿透率，abs是絕對值，取最大的，還有min、avg等規格
                            if (pointresult.Count() > 0)
                            {
                                //依每項分光規格來跑迴圈
                                foreach (var item1 in mtlQcitemresult)
                                {
                                    List<int> pointList = new List<int>();
                                    List<float> pointValueList = new List<float>();
                                    string design = item1.DesignValue.ToString();
                                    string upper = item1.UpperTolerance.ToString();
                                    string lower = item1.LowerTolerance.ToString();
                                    string[] wave = design.Split('~');
                                    string designType = upper.Split('.')[0].Replace("R", string.Empty).Replace("T", string.Empty);
                                    string designValue = lower.Replace("%", string.Empty);
                                    string judgeType = upper.Split('.')[1];

                                    //篩選點資料範圍
                                    foreach (var item2 in pointresult)
                                    {
                                        var point = item2.Point;
                                        var pointValue = item2.PointValue;

                                        if (wave.Count() > 1) //範圍 380~460
                                        {
                                            if (point >= Convert.ToInt16(wave[0]) && point <= Convert.ToInt16(wave[1]))
                                            {
                                                pointList.Add(Convert.ToInt16(point));
                                                pointValueList.Add(float.Parse(pointValue));
                                            }
                                        }
                                        else //單點 420
                                        {
                                            if (point != Convert.ToInt16(wave[0]))
                                            {
                                                break;
                                            }
                                            pointList.Add(Convert.ToInt16(point));
                                            pointValueList.Add(float.Parse(pointValue));
                                        }
                                    }

                                    //從點資料範圍中，找出規格類別所對應的值
                                    switch (designType)
                                    {
                                        case "abs":
                                            result.Add(pointValueList.Max().ToString());
                                            break;
                                        case "max":
                                            result.Add(pointValueList.Max().ToString());
                                            break;
                                        case "min":
                                            result.Add(pointValueList.Min().ToString());
                                            break;
                                        case "ave":
                                            result.Add(pointValueList.Average().ToString());
                                            break;
                                        default:
                                            throw new SystemException("查無 " + designType + " 規格類型，請洽系統開發室！");
                                    }

                                    foreach (var item3 in result)
                                    {
                                        if (decimal.TryParse(item3, out decimal s) == false) { continue; }

                                        //判定
                                        switch (judgeType)
                                        {
                                            case ">":
                                                if (float.Parse(item3) > float.Parse(designValue)) { judge.Add("OK"); } else { judge.Add("NG"); }
                                                break;
                                            case "<":
                                                if (float.Parse(item3) < float.Parse(designValue)) { judge.Add("OK"); } else { judge.Add("NG"); }
                                                break;
                                            case ">=":
                                                if (float.Parse(item3) >= float.Parse(designValue)) { judge.Add("OK"); } else { judge.Add("NG"); }
                                                break;
                                            case "<=":
                                                if (float.Parse(item3) <= float.Parse(designValue)) { judge.Add("OK"); } else { judge.Add("NG"); }
                                                break;
                                            case "=":
                                                if (float.Parse(item3) == float.Parse(designValue)) { judge.Add("OK"); } else { judge.Add("NG"); }
                                                break;
                                            default:
                                                throw new SystemException("查無 " + judgeType + " 判定類型，請洽系統開發室！");
                                        }
                                    }
                                }
                            }
                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = new { result, judge }
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
        #region //AddMeasureDataModifyFile -- 新增修改後量測檔案(共夾路徑) -- WuTc 2024-10-15
        public string AddMeasureDataModifyFile(string QcRecordId, string FilePath, string InputType)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, PhysicalPath
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @PhysicalPath
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcRecordId,
                                FileType = InputType,
                                PhysicalPath = FilePath,
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
        #endregion

        #region //Update

        #endregion

        #region //Delete

        #endregion

        #region //Upload Data
        #region //UploadModifyMeasureDataForMold -- 新增修改後上傳的excel數據(模仁) -- WuTc -- 2024-08-10
        public string UploadModifyMeasureDataForMold(List<UploadModifyDataModel> ModifyDataSetList, string QcRecordId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //先找出原始的量測數據ID
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT QmdId, QcItemId, BarcodeId 
		                    FROM QMS.QcMeasureData
	                        WHERE QcRecordId = @QcRecordId";

                    dynamicParameters.Add("QcRecordId", QcRecordId);
                    List<UploadModifyDataModel> resultMeasureDatas = sqlConnection.Query<UploadModifyDataModel>(sql, dynamicParameters).ToList();
                    List<UploadModifyDataModel> addModifyDatas = new List<UploadModifyDataModel>();

                    ModifyDataSetList = ModifyDataSetList.Join(resultMeasureDatas, x => new { x.QcItemId, x.BarcodeId }, y => new { y.QcItemId, y.BarcodeId }, (x, y) => { x.QmdId = y.QmdId; return x; }).ToList();
                    #endregion

                    int rowsAffected = 0;
                    ModifyDataSetList
                        .ToList()
                        .ForEach(x =>
                        {
                            x.CreateDate = CreateDate;
                            x.LastModifiedDate = LastModifiedDate;
                            x.CreateBy = CreateBy;
                            x.LastModifiedBy = LastModifiedBy;
                        });

                    dynamicParameters = new DynamicParameters();
                    sql = @"INSERT INTO QMS.QcMeasureDataModify (QmdId, QcRecordId, MoId, QcItemId, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, ZAxis, ModifyValue, BarcodeId, LetteringSeq, QmmDetailId, BallMark, Unit
                            , Surveyor, Confirmer, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
	                        VALUES (@QmdId, @QcRecordId, @MoId, @QcItemId, @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance, @ZAxis, @ModifyValue, @BarcodeId, @LetteringSeq, @QmmDetailId, @BallMark, @Unit
                            , @Surveyor, @Confirmer, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                    rowsAffected += sqlConnection.Execute(sql, ModifyDataSetList);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = rowsAffected
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
    }
}
