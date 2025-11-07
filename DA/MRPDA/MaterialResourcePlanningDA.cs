using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace MRPDA
{
    public class MaterialResourcePlanningDA
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
        private object demandLIneDtl;

        public MaterialResourcePlanningDA()
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
        #region //GetERPWipQty -- 取得ERP 未結製令數量 -- William 2023-01-31
        public int GetERPWipQty(int CompanyId, string MtlItemNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return 0;
        }
        #endregion

        #region //GetMtlItemQty
        private List<MtlItemQuantity> GetMtlItemQty()
        {
            List<MtlItemQuantity> mtlItemQuantity = new List<MtlItemQuantity>();

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT RTRIM(LTRIM(MB001)) MtlItemNo,SUM(ISNULL(MB064,0)) OnhandQty,SUM(ISNULL(b.OpenWipQty,0)) OpenWipQty, SUM(ISNULL(c.OpenPoQty,0)) OpenPoQty
                              FROM INVMB a 
                                   LEFT JOIN (SELECT TA006 MtlItemNo,SUM(TA015 - TA017) OpenWipQty
	                                            FROM MOCTA 
				                               WHERE TA011 NOT IN ('y','Y') 
				                                 AND TA013 NOT IN ('N','V')				                                     
								   		       GROUP BY TA006
											   HAVING SUM(TA015 - TA017) > 0) b ON a.MB001 = b.MtlItemNo
                                   LEFT JOIN (SELECT TD.TD004 MtlItemNo,SUM(TD.TD008 - TD.TD015) OpenPoQty
   			                                    FROM PURTC TC
					                                 LEFT JOIN PURTD TD ON TC.TC001 = TD.TD001 AND TC.TC002 = TD.TD002
				                               WHERE TC.TC014 = 'Y' 
				                                 AND TD016 = 'N' 					                                 
									  	       GROUP BY TD.TD004
											   HAVING SUM(TD.TD008 - TD.TD015) > 0) c ON a.MB001 = c.MtlItemNo
                            WHERE (a.MB031 > CONVERT(VARCHAR(8), GETDATE(), 112) OR a.MB031 = '')
                              AND ISNULL(MB064,0) + ISNULL(b.OpenWipQty,0) + ISNULL(c.OpenPoQty,0) > 0
							  GROUP BY RTRIM(LTRIM(MB001))";

                    mtlItemQuantity = sqlConnection.Query<MtlItemQuantity>(sql, dynamicParameters).ToList();
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return mtlItemQuantity;
        }
        #endregion

        #region //GetMtlItemUseQty
        private List<MtlItemUseQty> GetMtlItemUseQty()
        {
            List<MtlItemUseQty> mtlItemUseQty = new List<MtlItemUseQty>();

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DemandLineId,MtlItemNo,UsePoQty,UseWipQty,UseOnhandQty
                              FROM MRP.MtlItemUseQty a ";

                    mtlItemUseQty = sqlConnection.Query<MtlItemUseQty>(sql, dynamicParameters).ToList();
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return mtlItemUseQty;
        }
        #endregion

        #region //GetDemandLine
        private List<DemandLine> GetDemandLine(int DemandId)
        {
            List<DemandLine> demandLine = new List<DemandLine>();

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DemandId, a.DemandLineId,a.DemandPriority,a.SourceType
                                   ,a.SourceId, a.SourceNo, a.SourceSeq, a.MtlItemId, a.MtlItemNo, a.MakeType
                                   ,a.Quantity, a.CustomerId, a.ScheduleDate, a.ExpectDeliveryDate
                              FROM MRP.DemandLine a
                             WHERE a.DemandId = @DemandId";
                    dynamicParameters.Add("DemandId", DemandId);

                    demandLine = sqlConnection.Query<DemandLine>(sql, dynamicParameters).ToList();
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return demandLine;
        }
        #endregion

        #region //GetDemandLineDtl
        private List<DemandLineDtl> GetDemandLineDtl(int DemandId)
        {
            List<DemandLineDtl> demandLIneDtl = new List<DemandLineDtl>();

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @rowsAdded int
                            DECLARE @demandLIneDtl TABLE
                            ( 
                                DemandLineId int,
                                BomLevel int,
                                MtlItemNo nvarchar(30),
                                MtlItemId int,                                
                                ParentMtlItemId int,
                                MakeType nvarchar(2),
                                Quantity float,
                                CompositionQuantity float,
                                Base float,
                                ScheduleDate datetime,
                                WorkDays int,
                                processed int DEFAULT(0)
                            )

                            INSERT @demandLIneDtl
                                SELECT a.DemandLineId,
                                       1 BomLevel,
                                       LTRIM(RTRIM(b.MtlItemNo)) MtlItemNo,
                                       a.MtlItemId,
                                       a.MtlItemId ParentMtlItemId,
                                       b.ItemAttribute MakeType,
                                       a.Quantity,
                                       1 CompositionQuantity,
                                       1 Base,
                                       a.ScheduleDate,
                                       ISNULL(c.FixedLeadTime,0) + ISNULL(c.ChangeLeadTime,0) WorkDays,
                                       0 Processed
                                  FROM MRP.DemandLine a
                                       INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                       LEFT JOIN PDM.MtlItemSetting c ON b.MtlItemId = c.MtlItemId
                            WHERE a.DemandId = @DemandId

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @demandLIneDtl SET processed = 1 WHERE processed = 0

                                INSERT @demandLIneDtl
                                    SELECT a.DemandLineId,
                                           a.BomLevel + 1 BomLevel,
                                           LTRIM(RTRIM(d.MtlItemNo)) MtlItemNo,
                                           d.MtlItemId,
                                           b.MtlItemId ParentMtlItemId,
                                           d.ItemAttribute MakeType,
                                           a.Quantity / ISNULL(c.CompositionQuantity,1) * c.Base Quantity,                                           
                                           c.CompositionQuantity,
                                           c.Base,                                           
                                           a.ScheduleDate,
                                           ISNULL(e.FixedLeadTime,0) + ISNULL(e.ChangeLeadTime,0) WorkDays,
                                           0 Processed
                                      FROM @demandLIneDtl a
                                           INNER JOIN PDM.BillOfMaterials b ON a.MtlItemId = b.MtlItemId
                                           INNER JOIN PDM.BomDetail c ON b.BomId = c.BomId
                                           INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                                           LEFT JOIN PDM.MtlItemSetting e ON d.MtlItemId = e.MtlItemId
                                     WHERE a.processed = 1
                                       AND c.CompositionQuantity != 0

                                SET @rowsAdded = @@rowcount

                                UPDATE @demandLIneDtl SET processed = 2 WHERE processed = 1
                            END;

                            SELECT a.DemandLineId, a.BomLevel,a.MtlItemNo,a.MtlItemId,a.ParentMtlItemId,a.MakeType, a.Quantity,a.CompositionQuantity, a.Base,
                                   a.ScheduleDate,a.WorkDays                                   
                            FROM @demandLIneDtl a";
                    dynamicParameters.Add("DemandId", DemandId);

                    demandLIneDtl = sqlConnection.Query<DemandLineDtl>(sql, dynamicParameters).ToList();
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return demandLIneDtl;
        }
        #endregion
        #endregion

        #region //Add
        #region //AddDemand-- 新增主需求計劃MDS -- William 2023-01-17
        public string AddDemand(int CompanyId, string DemandName, string DemandDate, string EndDate, string DemandDesc)
        {
            try
            {
                if (DemandName.Length <= 0) throw new SystemException("【需求名稱】不能為空!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MRP.Demand (CompanyId, DemandName, DemandDate
                                , EndDate, DemandDesc
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DemandId
                                VALUES (@CompanyId, @DemandName, @DemandDate
                                , @EndDate, @DemandDesc
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                @DemandName,
                                @DemandDate,
                                @EndDate,
                                @DemandDesc,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //AddDemandLine-- 新增主需求明細(手動新增, 先預留, 理論上不會有機會觸發) -- William 2023-01-17
        public string AddDemandLine(int DemandId,int DemandPriority, string SourceType, int SourceId, string SourceNo, string SourceSeq,
            int MtlItemId, int Quantity, int CustomerId, DateTime ScheduleDate)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;

                        #region //同步SalesOrder裡的需求

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

        #region //AddDemandLineByBatch-- 新增主需求明細(由SalesOrder & ForecastOrder同步) -- William 2023-01-17
        public string AddDemandLineByBatch(int DemandId, string Status)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;

                        #region //匯入未結訂單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CustomerId,a.SoErpPrefix + '-' + a.SoErpNo OrderNo,
                                       c.MtlItemNo,c.MtlItemName,c.MtlItemDesc,a.Currency,
                                       b.SoQty,b.SiQty,b.SoQty - b.SiQty UnDeliveryQty,b.UnitPrice,
                                       b.SoSequence,b.SoDetailId,b.PromiseDate,c.MtlItemId,
                                       c.ItemAttribute MakeType
                                  FROM SCM.SaleOrder a
                                       INNER JOIN SCM.SoDetail b ON a.SoId = b.SoId
	                                   INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                 WHERE b.SoQty - b.SiQty > 0
                                   AND a.CompanyId = (SELECT CompanyId FROM MRP.Demand WHERE DemandId = @DemandId)
                                   AND b.ClosureStatus NOT IN ('y','Y')
                                   AND b.ConfirmStatus = 'Y'
                                   AND b.SoQty - b.SiQty > 0";

                        dynamicParameters.Add("DemandId", DemandId);
                        var OrderDemand = sqlConnection.Query(sql, dynamicParameters);
                        int DemandPriority = 1;
                        foreach (var item in OrderDemand)
                        {
                            int SoDetailId = Convert.ToInt32(item.SoDetailId);
                            string OrderNo = item.OrderNo;
                            int SoSequence = Convert.ToInt32(item.SoSequence);
                            int CustomerId = Convert.ToInt32(item.CustomerId);
                            int MtlItemId = Convert.ToInt32(item.MtlItemId);
                            string MtlItemNo = item.MtlItemNo;
                            string MakeType = item.MakeType;
                            double UnDeliveryQty = Convert.ToDouble(item.UnDeliveryQty);
                            DateTime PromiseDate;
                            if (item.PromiseDate == null)
                            {
                                PromiseDate = DateTime.Now;
                            }
                            else
                            {
                                PromiseDate = Convert.ToDateTime(item.PromiseDate);
                            }
                            

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MRP.DemandLine (DemandId, DemandPriority, SourceType
                                , SourceId, SourceNo, SourceSeq
                                , MtlItemId, MtlItemNo, MakeType, Quantity, CustomerId
                                , ScheduleDate
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DemandId
                                VALUES (@DemandId, @DemandPriority, @SourceType
                                , @SourceId, @SourceNo, @SourceSeq
                                , @MtlItemId, @MtlITemNo, @MakeType, @Quantity, @CustomerId
                                , @ScheduleDate
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DemandId,
                                    DemandPriority,
                                    SourceType = "SO",
                                    SourceId = SoDetailId,
                                    SourceNo = OrderNo,
                                    SourceSeq = SoSequence,
                                    MtlItemId,
                                    MtlItemNo,
                                    MakeType,
                                    Quantity = UnDeliveryQty,
                                    CustomerId,
                                    ScheduleDate = PromiseDate,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        }
                        #endregion

                        #region //匯入未轉訂單之需求, 只抓取有品號的需求

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
        #region //UpdateDemandStatus-- 更新DemanStatus(1:New,2:MRP需求已彙總,3:MRP計算完成) -- William 2023-01-17
        public string UpdateDemandStatus(int DemandId, string Status)
        {
            int WorkRow = 0;
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        switch (Status)
                        {
                            case "1":break;
                            case "2":
                                //呼叫AddDemandLineByBatch 新增DemandLine
                                AddDemandLineByBatch(DemandId, Status);
                                break;
                            case "3":
                                #region //先刪除舊有資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE MRP.DemandLineDtl
                                         WHERE DemandLineId IN(SELECT DemandLineId
                                                                 FROM MRP.DemandLine
                                                                WHERE DemandId = @DemandId)";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DemandId
                                });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE MRP.MtlItemUseQty
                                         WHERE DemandLineId IN(SELECT DemandLineId
                                                                 FROM MRP.DemandLine
                                                                WHERE DemandId = @DemandId)";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DemandId
                                });
                                var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //取得物料庫存相關資訊
                                List<MtlItemQuantity> mtlItemQuantity = new List<MtlItemQuantity>();
                                mtlItemQuantity = GetMtlItemQty();
                                #endregion

                                #region //取得DemandLine
                                List<DemandLineDtl> demandLineDtl = new List<DemandLineDtl>();
                                demandLineDtl = GetDemandLineDtl(DemandId);
                                #endregion

                                int MaxBomLevel = Convert.ToInt32(demandLineDtl.Max(x => x.BomLevel));

                                #region //由BomLevel低至高, 扣除現有供給數量更新至最下階
                                List<MtlItemUseQty> mtlItemUseQty = new List<MtlItemUseQty>();
                                int BomLevel = 1;
                                int demandLineCnt = demandLineDtl.Where(x => x.BomLevel == BomLevel).OrderBy(x => x.ScheduleDate).ToList().Count();

                                while (demandLineCnt > 0)
                                {
                                    demandLineDtl.Where(x => x.BomLevel == BomLevel).OrderBy(x => x.ScheduleDate).ToList().ForEach(x =>
                                    {
                                        #region //重新計算這階的需求量
                                        //DemandQty, 取上一階裡Quantity, 如果沒有則抓本階Quantity (Level1 沒有上階)                                         
                                        //抓出需求後, 重新*單位用量/底數
                                        double DemandQty = 0;
                                        if (BomLevel == 1)
                                        {
                                            DemandQty = Convert.ToDouble(x.Quantity);
                                        }
                                        else
                                        {
                                            DemandQty = Convert.ToDouble(demandLineDtl
                                                .Where(z => z.DemandLineId == x.DemandLineId && z.MtlItemId == x.ParentMtlItemId && z.BomLevel == BomLevel - 1)
                                                .Select(z => z.Quantity)
                                                .FirstOrDefault());
                                            //第一階以外需求量為上階需求量 * 單位用量 / 底數
                                            DemandQty = Convert.ToDouble(DemandQty * x.CompositionQuantity / x.Base);
                                        }

                                        #endregion
                                        double OpenPoQty = Convert.ToDouble(mtlItemQuantity.Where(y => y.MtlItemNo == x.MtlItemNo).Select(y => y.OpenPoQty).FirstOrDefault());
                                        double OpenWipQty = Convert.ToDouble(mtlItemQuantity.Where(y => y.MtlItemNo == x.MtlItemNo).Select(y => y.OpenWipQty).FirstOrDefault());
                                        double OnhandQty = Convert.ToDouble(mtlItemQuantity.Where(y => y.MtlItemNo == x.MtlItemNo).Select(y => y.OnhandQty).FirstOrDefault());
                                        double SumUsePoQty = Convert.ToDouble(mtlItemUseQty.Where(y => y.MtlItemNo == x.MtlItemNo).Select(y => y.UsePoQty).Sum());
                                        double SumUseWipQty = Convert.ToDouble(mtlItemUseQty.Where(y => y.MtlItemNo == x.MtlItemNo).Select(y => y.UseWipQty).Sum());
                                        double SumUseOnhandQty = Convert.ToDouble(mtlItemUseQty.Where(y => y.MtlItemNo == x.MtlItemNo).Select(y => y.UseOnhandQty).Sum());
                                        double UseOnhandQty = 0;
                                        double UsePoQty = 0;
                                        double UseWipQty = 0;
                                        #region //計算扣庫存與在途採購, 在製等數量取得實際需求
                                        if (x.MakeType.Equals("P"))
                                        {
                                            if (OnhandQty - SumUseOnhandQty > 0 && DemandQty > 0)
                                            {
                                                #region //先扣庫存
                                                if (OnhandQty - SumUseOnhandQty - DemandQty >= 0)
                                                {
                                                    UseOnhandQty = DemandQty;
                                                    DemandQty = 0;                                                    
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UseOnhandQty = UseOnhandQty });
                                                }
                                                else if (DemandQty - OnhandQty - SumUseOnhandQty > 0)
                                                {
                                                    UseOnhandQty = OnhandQty;
                                                    DemandQty = DemandQty - OnhandQty - SumUseOnhandQty;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UseOnhandQty = UseOnhandQty });
                                                }
                                                #endregion
                                            }
                                            if (OpenPoQty - SumUsePoQty > 0 && DemandQty > 0)
                                            {
                                                #region //再扣在途採購
                                                if (OpenPoQty - SumUsePoQty - DemandQty >= 0)
                                                {
                                                    UsePoQty = DemandQty;
                                                    DemandQty = 0;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UsePoQty = UsePoQty });
                                                }
                                                else if (DemandQty - OpenPoQty - SumUsePoQty > 0)
                                                {
                                                    UsePoQty = OpenPoQty;
                                                    DemandQty = DemandQty - OpenPoQty - SumUsePoQty;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UsePoQty = UsePoQty });
                                                }
                                                #endregion

                                            }
                                        }
                                        else if (x.MakeType.Equals("M"))
                                        {
                                            if (OnhandQty - SumUseOnhandQty > 0 && DemandQty > 0)
                                            {
                                                #region //先扣庫存
                                                if (OnhandQty - SumUseOnhandQty - DemandQty >= 0)
                                                {
                                                    UseOnhandQty = DemandQty;
                                                    DemandQty = 0;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UseOnhandQty = UseOnhandQty });
                                                }
                                                else if (DemandQty - OnhandQty - SumUseOnhandQty > 0)
                                                {
                                                    DemandQty = DemandQty - OnhandQty - SumUseOnhandQty;
                                                    UseOnhandQty = OnhandQty;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UseOnhandQty = UseOnhandQty });
                                                }
                                                #endregion
                                            }

                                            if (OpenWipQty - SumUseWipQty > 0 && DemandQty > 0)
                                            {
                                                #region //再扣在製
                                                if (OpenWipQty - SumUseWipQty - DemandQty >= 0)
                                                {
                                                    UseWipQty = DemandQty;
                                                    DemandQty = 0;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UseWipQty = UseWipQty });
                                                }
                                                else if (DemandQty - OpenWipQty - SumUseWipQty > 0)
                                                {
                                                    UseWipQty = OpenWipQty;
                                                    DemandQty = DemandQty - OpenWipQty - SumUseWipQty;
                                                    mtlItemUseQty.Add(new MtlItemUseQty() { DemandLineId = x.DemandLineId, MtlItemNo = x.MtlItemNo, UseWipQty = UseWipQty });
                                                }
                                                #endregion
                                            }
                                        }
                                        #endregion

                                        x.Quantity = DemandQty;
                                    });
                                    BomLevel++;
                                    demandLineCnt = demandLineDtl.Where(x => x.BomLevel == BomLevel).ToList().Count();

                                }

                                #region //寫入MRP.MtlItemUseQty & MRP.DemandLineDtl

                                if (mtlItemUseQty.Count > 0)
                                {
                                    sql = @"INSERT INTO MRP.MtlItemUseQty(DemandLineId, MtlItemNo, UsePoQty, UseWipQty, UseOnhandQty)
                                            VALUES(@DemandLineId, @MtlItemNo, @UsePoQty, @UseWipQty, @UseOnhandQty)";
                                    rowsAffected += sqlConnection.Execute(sql, mtlItemUseQty);
                                }
                                #endregion

                                #endregion

                                #region //先將ScheduleDate歸0, 由系統時間+2小時為開工或採購時間
                                demandLineDtl.ToList().ForEach(x =>
                                {
                                    if(x.ScheduleDate < Convert.ToDateTime(DateTime.Now.AddHours(2)))
                                    {
                                        x.ScheduleDate = Convert.ToDateTime(DateTime.Now.AddHours(2));
                                    }
                                });
                                #endregion

                                #region //找出MRP裡最大BomLevel, 由大到小更新ScheduleDate
                                demandLineCnt = demandLineDtl.Where(x => x.BomLevel == MaxBomLevel).ToList().Count();
                                while(demandLineCnt > 0)
                                {
                                    demandLineDtl.Where(x=>x.BomLevel == MaxBomLevel).ToList().ForEach(x=>{
                                        demandLineDtl.Where(y => y.BomLevel == MaxBomLevel - 1 
                                        && y.DemandLineId == x.DemandLineId
                                        && y.MtlItemId == x.ParentMtlItemId
                                        && x.Quantity != 0).ToList().ForEach(y =>
                                          {
                                              DateTime NewScheduleDate = Convert.ToDateTime(x.ScheduleDate);
                                              Double NewWorkDays = Convert.ToDouble(x.WorkDays);
                                              y.ScheduleDate = NewScheduleDate.AddHours(NewWorkDays);
                                          });
                                    });
                                    MaxBomLevel--;
                                    demandLineCnt = demandLineDtl.Where(x => x.BomLevel == MaxBomLevel).ToList().Count();
                                }
                                #endregion

                                #region //寫入MRP.DemandLineDtl
                                if (demandLineDtl.Count > 0)
                                {
                                    demandLineDtl
                                        .ToList()
                                        .ForEach(x =>
                                        {
                                            x.CreateDate = CreateDate;
                                            x.LastModifiedDate = LastModifiedDate;
                                            x.CreateBy = CreateBy;
                                            x.LastModifiedBy = LastModifiedBy;
                                        });

                                    sql = @"INSERT INTO MRP.DemandLineDtl(DemandLineId, BomLevel, MtlItemNo, MtlItemId, ParentMtlItemId, MakeType
                                            , Quantity, CompositionQuantity, Base, ScheduleDate
                                            , WorkDays, ExpectFinishDate, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES(@DemandLineId,@BomLevel,@MtlItemNo,@MtlItemId,@ParentMtlItemId,@MakeType
                                            ,@Quantity,@CompositionQuantity,@Base,@ScheduleDate
                                            ,@WorkDays,@ExpectFinishDate,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                                    rowsAffected += sqlConnection.Execute(sql, demandLineDtl);
                                }
                                #endregion

                                break;
                            default: throw new SystemException("Demand Status Error!!!");
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
                    status = "errorForDA" + " WorkRow=" + WorkRow,
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateDemandSchedule, 找出需要更新之DemandLineDtl呼叫UpdateDemandSchedule更新需求計劃明細之計劃開始日期
        private string UpdateDemandSchedule(int DemandId, SqlConnection sqlConnection)
        {
            try
            {
                int rowsAffected = 0;

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "(" + rowsAffected + " rows affected)"
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

        #region //UpdateDemandScheduleDate, 遞迴更新DemandLineDtl的ScheduleDate
        private void UpdateDemandScheduleDate(int DemandLineId, int BomLevel, SqlConnection sqlConnection)
        {

        }
        #endregion
        #endregion

        #region //Delete
        #endregion

        #region Function

        #endregion
    }
}
