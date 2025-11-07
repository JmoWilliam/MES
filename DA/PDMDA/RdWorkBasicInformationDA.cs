using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Dynamic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace PDMDA
{
    public class RdWorkBasicInformationDA
    {
        public static string MainConnectionStrings = "";
        public static string ErpConnectionStrings = "";
        public static string HrmConnectionStrings = "";

        public static int CurrentCompany = -1;
        public static int CurrentUser = -1;
        public static int CreateBy = -1;
        public static int LastModifiedBy = -1;
        public static string UserLock = "";
        public static string UserNo = "";
        public static DateTime CreateDate = default(DateTime);
        public static DateTime LastModifiedDate = default(DateTime);

        public static string sql = "";
        public static JObject jsonResponse = new JObject();
        public static Logger logger = LogManager.GetCurrentClassLogger();
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static SqlQuery sqlQuery = new SqlQuery();

        public RdWorkBasicInformationDA()
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
        #region //GetMtlItemPageList
        private List<dynamic> MtlItemPageList(int BomId, int MtlItemId, string MtlItemNo, string MtlItemName, string MtlItemSpec, string CustomerMtlItemNo
            , string Status, string TransferStatus, string BomTransferStatus, string CustomerTransferStatus, string BomSubstitutionTransferStatus, string UomCTransferStatus, string CheckMtlDate
            , int ItemTypeId, string LotManagement, string ItemAttribute
            , string OrderBy, int PageIndex, int PageSize)
        {
            using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
            {
                #region //篩選資料
                dynamicParameters = new DynamicParameters();
                sqlQuery.mainKey = "a.MtlItemId";
                sqlQuery.auxKey = "";
                sqlQuery.columns = ", a.MtlItemName";
                sqlQuery.mainTables =
                    @" FROM PDM.MtlItem a";
                sqlQuery.auxTables = "";
                sqlQuery.distinct = true;
                sql = " AND a.CompanyId = @CompanyId";

                dynamicParameters.Add("CompanyId", CurrentCompany);

                if (MtlItemId > 0)
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                }

                if (!string.IsNullOrWhiteSpace(TransferStatus))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TransferStatus", @" AND a.TransferStatus IN @TransferStatus", TransferStatus.Split(','));
                }

                if (!string.IsNullOrWhiteSpace(MtlItemNo))
                {
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                    //    , "MtlItemNo"
                    //    , @" AND a.MtlItemNo LIKE N'%' + @MtlItemNo + '%'", MtlItemNo.Trim());
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo = @MtlItemNo", MtlItemNo.Trim());
                }

                if (!string.IsNullOrWhiteSpace(MtlItemName))
                {
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                    //    , "MtlItemName"
                    //    , @" AND a.MtlItemName LIKE N'%' + @MtlItemName + '%'", MtlItemName.Trim());
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND a.MtlItemName = @MtlItemName", MtlItemName.Trim());
                }

                if (!string.IsNullOrWhiteSpace(MtlItemSpec))
                {
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters
                        , "MtlItemSpec"
                        , @" AND a.MtlItemSpec LIKE N'%' + @MtlItemSpec + '%'", MtlItemSpec.Trim());
                }

                sqlQuery.conditions = sql;
                sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MtlItemId";
                sqlQuery.pageIndex = PageIndex;
                sqlQuery.pageSize = PageSize;

                return BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();
                #endregion
            }
        }
        #endregion

        #region //GetBomRecursionString
        public string GetBomRecursionString(ref DynamicParameters dynamicParameters, List<dynamic> MtlItemIds)
        {
            var recursionString = "";
            recursionString = @"
                DECLARE @dtNow DATETIME = GETDATE()
                DECLARE @rowsAdded INT
                DECLARE @tempData TABLE
                (
	                MtlItemId INT,
                    ParentBomId INT,
	                ParentMtlItemId INT,
                    ParentMtlItemNo NVARCHAR(40),
	                BomId INT,
	                BomSequence VARCHAR(4),
	                CompositionQuantity FLOAT,
	                Base FLOAT, 
	                LossRate FLOAT,
	                MtlItemRoute VARCHAR(256),
                    lvl INT,
	                processed INT DEFAULT(0),
	                ErrorFlag INT
                )
                INSERT @tempData
	                SELECT a.MtlItemId
	                , -1
                    , -1
                    , ''
	                , ISNULL(b.BomId, -1)
	                , '0000'
	                , 1
	                , 1
	                , 0
	                , '-1'
	                , 1
	                , 0
                    , 0
	                FROM PDM.MtlItem a
	                LEFT JOIN PDM.BillOfMaterials b ON b.MtlItemId = a.MtlItemId
	                WHERE a.MtlItemId IN @MtlItemIds";

            dynamicParameters.Add("MtlItemIds", MtlItemIds);

            recursionString += @"
	            DECLARE @checkError INT = 0
	            DECLARE @tempAdd TABLE (CurrentMtlId INT, CurrentParentMtlId INT)
                SET @rowsAdded=@@rowcount
                DECLARE @maxlvl INT = 1
                WHILE @rowsAdded > 0
                BEGIN
                    UPDATE @tempData SET processed = 1 WHERE processed = 0	
                    INSERT @tempData
                        OUTPUT INSERTED.MtlItemId,INSERTED.ParentMtlItemId INTO @tempAdd
		                SELECT c.MtlItemId, a.BomId, a.MtlItemId, p.MtlItemNo, d.BomId
		                , b.BomSequence, b.CompositionQuantity, b.Base, b.LossRate
		                , CAST(a.MtlItemRoute + ',' + CAST(a.MtlItemId AS VARCHAR(256)) AS VARCHAR(256))
		                , (a.lvl + 1)
		                , 0
                        , 0
		                FROM @tempData a
                        INNER JOIN PDM.MtlItem p ON p.MtlItemId = a.MtlItemId
			            INNER JOIN PDM.BomDetail b ON b.BomId = a.BomId
			            INNER JOIN PDM.MtlItem c ON c.MtlItemId = b.MtlItemId
			            LEFT JOIN PDM.BillOfMaterials d ON d.MtlItemId = c.MtlItemId
			            INNER JOIN PDM.UnitOfMeasure e ON e.UomId = b.UomId
		                WHERE a.processed = 1
			            AND @dtNow >= ISNULL(b.EffectiveDate, '1900-01-01')
			            AND @dtNow <= ISNULL(b.ExpirationDate, '9999-01-01')
                    SET @rowsAdded = @@rowcount

		            UPDATE a SET a.ErrorFlag = 1, @checkError = 1 
		            FROM @tempData a
		            INNER JOIN @tempAdd b ON b.CurrentMtlId = a.ParentMtlItemId AND b.CurrentParentMtlId = a.MtlItemId
		            IF (@checkError = 1) 
		            BEGIN
			            SET @rowsAdded = 0
		            END

                    UPDATE @tempData SET processed = 2 WHERE processed = 1
                END
            ";

            return recursionString;
        }
        #endregion

        #region //GetBillOfMaterialsTree -- 取得Bom(樹狀資料)(非結構) -- Chia Yuan 2023.8.21
        public string GetBillOfMaterialsTree(int BomId, int MtlItemId, string MtlItemNo, string MtlItemName, string MtlItemSpec, string CustomerMtlItemNo
            , string Status, string TransferStatus, string BomTransferStatus, string CustomerTransferStatus, string BomSubstitutionTransferStatus, string UomCTransferStatus, string CheckMtlDate
            , int ItemTypeId, string LotManagement, string ItemAttribute)
        {
            try
            {
                var mtlItems = MtlItemPageList(BomId, MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, CustomerMtlItemNo
                    , Status, TransferStatus, BomTransferStatus, CustomerTransferStatus, BomSubstitutionTransferStatus, UomCTransferStatus, CheckMtlDate
                    , ItemTypeId, LotManagement, ItemAttribute
                    , "", -1, -1);

                var mtlItemIds = mtlItems.Select(s => s.MtlItemId).ToList();
                List<BillOfMaterialsTree> bomTree = new List<BillOfMaterialsTree>();
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //樹狀資料
                    dynamicParameters = new DynamicParameters();
                    sql = GetBomRecursionString(ref dynamicParameters, mtlItemIds);
                    sql += @"SELECT *, ROW_NUMBER() OVER (PARTITION BY a.ParentMtlItemId ORDER BY a.BomSequence) AS SortNumber
                            FROM (
                                SELECT a.*
                                , b.CompanyId, LTRIM(RTRIM(b.MtlItemNo)) AS MtlItemNo, LTRIM(RTRIM(b.MtlItemName)) AS MtlItemName, LTRIM(RTRIM(b.MtlItemSpec)) AS MtlItemSpec
                                , b.ItemTypeId, b.InventoryId, b.RequisitionInventoryId, b.InventoryManagement
                                , b.PurchaseUomId, b.SaleUomId, b.MtlModify, b.BondedStore, b.ItemAttribute, b.ProductionLine, b.LotManagement, b.MeasureType, b.InventoryUomId
                                , b.EffectiveDate, b.ExpirationDate, b.OverReceiptManagement, b.OverDeliveryManagement, b.Version, b.MtlItemDesc, b.MtlItemRemark
                                , b.TransferStatus, b.TransferDate, b.EfficientDays, b.RetestDays, b.ReplenishmentPolicy
                                , b.Status, s1.StatusName
                                , b.CreateDate
                                , CASE WHEN GETDATE() BETWEEN ISNULL(b.EffectiveDate, '1900-01-01') AND ISNULL(b.ExpirationDate, '9999-01-01') THEN 1 ELSE 0 END AS Effective
                                , ISNULL(c.UomNo, '') AS InventoryUomNo
                                , ISNULL(u1.UserName, '') AS UserName
                                , ISNULL(bom.BomTransferQty, 0) AS BomTransferQty
                                , ISNULL(Cus.CusTransferQty, 0) AS CusTransferQty
                                , ISNULL(Sub.SubTransferQty, 0) AS SubTransferQty
                                , ISNULL(Uom.UomTransferQty, 0) AS UomTransferQty
                                , ISNULL(b.TypeOne, '') AS TypeOne
                                , ISNULL(b.TypeTwo, '') AS TypeTwo
                                , ISNULL(b.TypeThree, '') AS TypeThree
                                , ISNULL(b.TypeFour, '') AS TypeFour
                                , ISNULL(i.InventoryNo, '') AS InventoryNo, ISNULL(i.InventoryName, '') AS InventoryName, i.InventoryNo + ' ' + i.InventoryName AS InventoryWithNo
                                , ISNULL(d.ItemTypeNo, '') AS ItemTypeNo, ISNULL(d.ItemTypeName, '') AS ItemTypeName, d.ItemTypeNo + ' ' + d.ItemTypeName AS ItemWithNo
	                            FROM @tempData a
	                            INNER JOIN PDM.MtlItem b ON b.MtlItemId = a.MtlItemId
                                LEFT JOIN PDM.UnitOfMeasure c on c.UomId = b.InventoryUomId
                                LEFT JOIN PDM.ItemType d ON d.ItemTypeId = b.ItemTypeId
                                INNER JOIN SCM.Inventory i ON i.InventoryId = b.InventoryId
                                OUTER APPLY(
	                                SELECT COUNT(x1.MtlItemId) AS BomTransferQty FROM PDM.BillOfMaterials x1 INNER JOIN PDM.BomDetail x2 ON x2.BomId = x1.BomId 
                                    WHERE x1.MtlItemId = b.MtlItemId AND x1.[TransferStatus] = 'N'
                                ) bom
                                OUTER APPLY(
	                                SELECT COUNT(x1.MtlItemId) AS CusTransferQty FROM PDM.CustomerMtlItem x1 
                                    WHERE x1.MtlItemId = b.MtlItemId AND x1.[TransferStatus] = 'N'
                                ) Cus
                                OUTER APPLY(
	                                SELECT COUNT(x1.MtlItemId) AS SubTransferQty FROM PDM.BomSubstitution x1 INNER JOIN PDM.BomDetail x2 ON x2.BomDetailId = x1.BomDetailId INNER JOIN PDM.BillOfMaterials x3 ON x3.BomId = x2.BomId 
                                    WHERE x3.MtlItemId = b.MtlItemId AND x1.[TransferStatus] = 'N'
                                ) Sub
                                OUTER APPLY(
	                                SELECT COUNT(x1.MtlItemId) AS UomTransferQty FROM PDM.MtlItemUomCalculate x1 
                                    WHERE x1.MtlItemId = b.MtlItemId AND x1.[TransferStatus] = 'N'
                                ) Uom
                                INNER JOIN BAS.[Status] s1 ON s1.StatusNo = b.Status AND s1.StatusSchema = 'Status'
                                LEFT JOIN BAS.[User] u1 on u1.UserId = b.CreateBy
                            ) a
                            ORDER BY a.lvl, a.BomSequence";
                    bomTree = sqlConnection.Query<BillOfMaterialsTree>(sql, dynamicParameters).ToList();
                    #endregion

                    var e = bomTree.FirstOrDefault(f => f.ErrorFlag == 1);
                    if (e != null)
                    {
                        throw new SystemException(string.Format(@"品號:{0}<=>{1}，主元件循環參考!", e.MtlItemNo, e.ParentMtlItemNo));
                    }

                    #region //品號編碼節段
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ItemTypeId, a.ItemTypeSegmentId, a.SortNumber, a.SegmentType, a.SegmentValue, a.SuffixCode 
                            , ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.InactiveDate, 'yyyy-MM-dd'), '') InactiveDate
                            FROM PDM.ItemTypeSegment a 
                            WHERE a.ItemTypeId IN @ItemTypeIds 
                            ORDER BY a.ItemTypeId, a.SortNumber";
                    dynamicParameters.Add("ItemTypeIds", bomTree.Where(w => w.ItemTypeId > 0).Select(s => s.ItemTypeId).Distinct().ToArray());
                    var resultItemTypeSegment = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    bomTree.Where(w => w.ItemTypeId > 0).ToList()
                        .ForEach(x =>
                        {
                            var itemTypeSegments = resultItemTypeSegment.Where(w => Convert.ToInt32(w.ItemTypeId) == x.ItemTypeId.Value).ToList();
                            if (itemTypeSegments.Any())
                            {
                                x.ItemTypeSegments = itemTypeSegments.Select(s => new ItemTypeSegment {
                                    ItemTypeSegmentId = s.ItemTypeSegmentId,
                                    SortNumber = s.SortNumber,
                                    SegmentType = s.SegmentType,
                                    SegmentValue = s.SegmentValue,
                                    SuffixCode = s.SuffixCode,
                                    EffectiveDate = s.EffectiveDate,
                                    InactiveDate = s.InactiveDate
                                }).ToList(); //取得編碼節段資料
                            }
                        });

                    #region //取得品號公司別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId IN @CompanyIds";
                    dynamicParameters.Add("CompanyIds", bomTree.Select(s => s.CompanyId).Distinct().ToArray());

                    var resultCorp = sqlConnection.Query(sql, dynamicParameters);
                    foreach (var corp in resultCorp)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[corp.ErpDb];
                    }
                    #endregion
                }

                if (!string.IsNullOrWhiteSpace(ErpConnectionStrings))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        string[] mtlItemNos = bomTree.Select(s => s.MtlItemNo).ToArray();

                        #region //撈取類別一、二、三、四
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MA002)) AS MtlItemTypeNo
                                , LTRIM(RTRIM(MA003)) AS MtlItemTypeName
                                FROM INVMA
                                WHERE 1=1";

                        List<Erp> erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();
                        bomTree = bomTree.GroupJoin(erps, x => x.TypeOne, y => y.MtlItemTypeNo, (x, y) => { x.TypeOneName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                        bomTree = bomTree.GroupJoin(erps, x => x.TypeTwo, y => y.MtlItemTypeNo, (x, y) => { x.TypeTwoName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                        bomTree = bomTree.GroupJoin(erps, x => x.TypeThree, y => y.MtlItemTypeNo, (x, y) => { x.TypeThreeName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                        bomTree = bomTree.GroupJoin(erps, x => x.TypeFour, y => y.MtlItemTypeNo, (x, y) => { x.TypeFourName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                        #endregion

                        #region //撈取品號總庫存量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MC001, ISNULL(SUM(MC007), 0) AS TotalQty
                                FROM INVMC
                                WHERE MC001 IN @MC001s
                                GROUP BY MC001";
                        dynamicParameters.Add("MC001s", mtlItemNos);
                        var resultINVMC = sqlConnection.Query(sql, dynamicParameters).ToList();

                        bomTree.ForEach(item =>
                        {
                            item.IsNewItem = mtlItemIds.Contains(item.MtlItemId) ? true : false;
                            item.IsClone = true;
                            item.TotalInventoryQty = resultINVMC.FirstOrDefault(f => f.MC001 == item.MtlItemNo)?.TotalQty ?? 0;
                        });
                        #endregion
                    }
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = bomTree
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

        #region //GetMtlItem -- 取得品號資料 -- Ben Ma 2022.06.30
        public string GetMtlItem(int MtlItemId, string MtlItemNo, string MtlItemName, string MtlItemSpec, string CustomerMtlItemNo
            , string TransferStatus, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<MtlItem> mtlItems = new List<MtlItem>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤:" + UserNo + ":" + UserLock + ":" + CurrentUser + ":" + CurrentCompany);

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    sqlQuery.mainKey = "a.MtlItemId";
                    sqlQuery.auxKey = ", a.CustomerMtlItem";
                    sqlQuery.columns =
                        @", a.ItemTypeId, a.MtlItemNo, a.MtlItemName, a.MtlItemSpec, a.PurchaseUomId, a.SaleUomId,a.TransferStatus
                          , ISNULL(a.TypeOne, '') TypeOne, ISNULL(a.TypeTwo, '') TypeTwo, ISNULL(a.TypeThree, '') TypeThree
                          , ISNULL(a.TypeFour, '') TypeFour, a.InventoryId, a.RequisitionInventoryId, a.InventoryManagement
                          , a.MtlModify, a.BondedStore, a.ItemAttribute, a.LotManagement, a.MeasureType, a.InventoryUomId
                          , a.EffectiveDate, a.ExpirationDate, a.Version, a.MtlItemDesc, a.MtlItemRemark
                          , a.TransferStatus, a.TransferDate, a.Status, a.CreateDate
                          , ISNULL(b.UomNo, '') AS InventoryUomNo
                          , c.UserName
                          , x.MtlItemUomCalculateQty";
                    sqlQuery.mainTables =
                        @"FROM PDM.MtlItem a
                          LEFT JOIN PDM.UnitOfMeasure b on b.UomId = a.InventoryUomId
                          LEFT JOIN BAS.[User] c on c.UserId = a.CreateBy
                          OUTER APPLY(
                            SELECT COUNT(x1.MtlItemUomCalculateId) MtlItemUomCalculateQty  FROM PDM.MtlItemUomCalculate x1 WHERE a.MtlItemId = x1.MtlItemId
                            ) x";
                    string queryTable =
                        @"FROM (
                            SELECT a.MtlItemId, a.CompanyId, a.ItemTypeId, a.MtlItemNo, a.MtlItemName, a.MtlItemSpec, a.TransferStatus,c.UserName, x.MtlItemUomCalculateQty
                            , (
                                    SELECT aa.CustomerMtlItemId, aa.CustomerId, aa.CustomerMtlItemNo, aa.CustomerMtlItemName, aa.CustomerMtlItemSpec
                                    , ab.CustomerNo, ab.CustomerShortName, aa.CustomerId customerId, aa.CustomerMtlItemNo customerMtlItemNo
                                    , aa.CustomerMtlItemName customerMtlItemName, aa.CustomerMtlItemSpec customerMtlItemSpec
                                    , ab.CustomerNo + ' ' + ab.CustomerShortName customerName
                                    FROM PDM.CustomerMtlItem aa
                                    INNER JOIN SCM.Customer ab ON aa.CustomerId = ab.CustomerId
                                    WHERE aa.MtlItemId = a.MtlItemId
                                    ORDER BY ab.CustomerNo, aa.CustomerMtlItemNo
                                    FOR JSON PATH, ROOT('data')
                            ) CustomerMtlItem
                            FROM PDM.MtlItem a
                            LEFT JOIN PDM.UnitOfMeasure b on a.InventoryUomId = b.UomId
                            LEFT JOIN BAS.[User] c on a.CreateBy = c.UserId
                            OUTER APPLY(
                                SELECT COUNT(x1.MtlItemUomCalculateId) MtlItemUomCalculateQty  FROM PDM.MtlItemUomCalculate x1 WHERE a.MtlItemId = x1.MtlItemId
                            ) x
                        ) a 
                        OUTER APPLY (
                            SELECT TOP 1 x.CustomerMtlItemNo, x.CustomerMtlItemName, x.CustomerMtlItemSpec
                            FROM OPENJSON(a.CustomerMtlItem, '$.data')
                            WITH (
                                CustomerMtlItemNo NVARCHAR(50) N'$.CustomerMtlItemNo',
                                CustomerMtlItemName NVARCHAR(200) N'$.CustomerMtlItemName',
                                CustomerMtlItemSpec NVARCHAR(200) N'$.CustomerMtlItemSpec'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND a.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemSpec", @" AND a.MtlItemSpec LIKE '%' + @MtlItemSpec + '%'", MtlItemSpec);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND b.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransferStatus", @" AND a.TransferStatus = @TransferStatus", TransferStatus);
                    if (SearchKey.Length > 0)
                    {
                        string searchSql = "";
                        string[] searchKey = SearchKey.Split(' ');

                        searchSql += @" AND 
                                        (
                                            b.CustomerMtlItemNo LIKE '%' + @SearchKey + '%'
                                            OR a.MtlItemName LIKE '%' + @SearchKey + '%'
                                            OR a.MtlItemNo LIKE '%' + @SearchKey + '%'";
                        if (searchKey[0].Length >= 4)
                        {
                            searchSql += @" OR b.CustomerMtlItemNo LIKE '%' + @ParagraphStep1 + '%'
                                            OR a.MtlItemName LIKE '%' + @ParagraphStep1 + '%'";
                            if (searchKey[0].Length >= 6)
                            {
                                searchSql += @" OR b.CustomerMtlItemNo LIKE '%' + @ParagraphStep2 + '%'
                                                OR a.MtlItemName LIKE '%' + @ParagraphStep2 + '%'
                                                OR a.MtlItemNo LIKE '%' + @ParagraphStep2 + '%'";
                            }

                            dynamicParameters.Add("ParagraphStep1", searchKey[0]);
                            dynamicParameters.Add("ParagraphStep2", searchKey[0].Substring(2, searchKey[0].Length - 2));
                        }
                        if (searchKey.Length > 1)
                        {
                            if (searchKey[1].Length >= 4)
                            {
                                searchSql += @" OR b.CustomerMtlItemNo LIKE '%' + @ParagraphStep3 + '%'
                                                OR a.MtlItemName LIKE '%' + @ParagraphStep3 + '%'
                                                OR a.MtlItemNo LIKE '%' + @ParagraphStep3 + '%'";
                                if (searchKey[1].Length >= 6)
                                {
                                    searchSql += @" OR b.CustomerMtlItemNo LIKE '%' + @ParagraphStep4 + '%'
                                                    OR a.MtlItemName LIKE '%' + @ParagraphStep4 + '%'
                                                    OR a.MtlItemNo LIKE '%' + @ParagraphStep4 + '%'";
                                }
                                dynamicParameters.Add("ParagraphStep3", searchKey[1]);
                                dynamicParameters.Add("ParagraphStep4", searchKey[1].Substring(2, searchKey[1].Length - 2));
                            }
                        }
                        searchSql += ")";
                        dynamicParameters.Add("SearchKey", SearchKey);

                        queryCondition += searchSql;
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MtlItemId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    mtlItems = BaseHelper.SqlQuery<MtlItem>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<Erp> erps = new List<Erp>();

                    #region //撈取類別一、二、三、四
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MA002)) MtlItemTypeNo, LTRIM(RTRIM(MA003)) MtlItemTypeName
                            FROM INVMA
                            WHERE 1=1";
                    erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                    mtlItems = mtlItems.GroupJoin(erps, x => x.TypeOne, y => y.MtlItemTypeNo, (x, y) => { x.TypeOneName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                    mtlItems = mtlItems.GroupJoin(erps, x => x.TypeTwo, y => y.MtlItemTypeNo, (x, y) => { x.TypeTwoName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                    mtlItems = mtlItems.GroupJoin(erps, x => x.TypeThree, y => y.MtlItemTypeNo, (x, y) => { x.TypeThreeName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                    mtlItems = mtlItems.GroupJoin(erps, x => x.TypeFour, y => y.MtlItemTypeNo, (x, y) => { x.TypeFourName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; }).ToList();
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = mtlItems
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

        #region //GetCustomerMtlItem -- 取得客戶品號資料 -- Zoey 2022.07.14
        public string GetCustomerMtlItem(int CustomerMtlItemId, int MtlItemId, int CustomerId, string CustomerMtlItemNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CustomerMtlItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MtlItemId, a.CustomerId, a.CustomerMtlItemNo, a.CustomerMtlItemName, a.CustomerMtlItemSpec
                          , b.CustomerNo, b.CustomerShortName, b.CustomerName
                          , c.MtlItemNo, c.MtlItemName, c.MtlItemSpec, c.TransferStatus";
                    sqlQuery.mainTables =
                        @"FROM PDM.CustomerMtlItem a
                          INNER JOIN SCM.Customer b ON b.CustomerId = a.CustomerId
                          INNER JOIN PDM.MtlItem c ON c.MtlItemId = a.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemId", @" AND a.CustomerMtlItemId = @CustomerMtlItemId", CustomerMtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND a.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CustomerMtlItemId";
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

        #region //GetBillOfMaterials -- 取得Bom主件資料 -- Zoey 2022.12.05
        public string GetBillOfMaterials(int BomId, int MtlItemId)
        {
            try
            {
                List<BillOfMaterials> boms = new List<BillOfMaterials>();

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

                    sql = @"SELECT a.BomId, a.MtlItemId, c.MtlItemNo, c.MtlItemName, c.MtlItemSpec
                            , a.UomId, a.StandardLot, a.WipPrefix, a.ModiPrefix, a.ModiNo, a.ModiSequence, a.Version
                            , a.Remark, a.ConfirmStatus, a.ConfirmUserId, a.ConfirmDate, a.TransferStatus, a.TransferDate, a.Status
                            , b.UomNo
                            FROM PDM.BillOfMaterials a
                            INNER JOIN PDM.UnitOfMeasure b ON a.UomId = b.UomId
                            INNER JOIN PDM.MtlItem c ON c.MtlItemId = a.MtlItemId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BomId", @" AND a.BomId = @BomId", BomId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);

                    boms = sqlConnection.Query<BillOfMaterials>(sql, dynamicParameters).ToList();
                }
                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<Erp> erps = new List<Erp>();

                    #region //撈取製令單別
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT LTRIM(RTRIM(MQ001)) WipPrefix, LTRIM(RTRIM(MQ002)) WipPrefixName 
                            FROM CMSMQ 
                            LEFT JOIN CMSMU ON MQ001 = MU001 
                            WHERE (MQ003 = '51' OR MQ003 = '52')  
                            AND ((MU003 = 'DS' AND MQ029 = 'Y')  OR MQ029 = 'N' )";
                    erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                    boms = boms.GroupJoin(erps, x => x.WipPrefix, y => y.WipPrefix, (x, y) => { x.WipPrefixName = y.FirstOrDefault()?.WipPrefixName ?? ""; return x; }).ToList();
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = boms
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

        #region //GetBomDetail -- 取得Bom元件資料 -- Zoey 2022.12.05
        public string GetBomDetail(int BomDetailId, int BomId, int MtlItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.BomDetailId, a.BomId, a.BomSequence, a.MtlItemId, a.UomId, a.CompositionQuantity, a.Base, a.LossRate,a.StandardCostingType,a.MaterialProperties
                            , ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, FORMAT(a.ExpirationDate, 'yyyy-MM-dd') ExpirationDate
                            , a.ReleaseItem, a.Remark, a.SubstitutionRemark, a.ConfirmStatus, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                            , b.TransferStatus
                            , c.MtlItemNo, c.MtlItemName, c.MtlItemSpec
                            , d.UomNo
                            FROM PDM.BomDetail a
                            INNER JOIN PDM.BillOfMaterials b ON b.BomId = a.BomId
                            INNER JOIN PDM.MtlItem c ON c.MtlItemId = a.MtlItemId
                            INNER JOIN PDM.UnitOfMeasure d ON a.UomId = d.UomId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BomDetailId", @" AND a.BomDetailId = @BomDetailId", BomDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BomId", @" AND a.BomId = @BomId", BomId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemId", @" AND b.MtlItemId  = @MtlItemId", MtlItemId);

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

        #region //CheckExistsMtlItemNo
        public string CheckExistsMtlItemNo(string BomStructureJson)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(BomStructureJson)) throw new SystemException("【結構】資料錯誤!");
                if (!BomStructureJson.TryParseJson(out JObject bomStructureJson)) throw new SystemException("【結構】資料錯誤!");
                if (bomStructureJson["data"].Count() <= 0) throw new SystemException("【結構】資料不存在!");

                List<BillOfMaterialsTree> newBomList = JsonConvert.DeserializeObject<List<BillOfMaterialsTree>>(JObject.Parse(BomStructureJson)["data"].ToString()).ToList(); //反序列化
                if (!newBomList.Any()) throw new SystemException("【結構】資料不存在!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string[] newMtlItemNos = newBomList.Where(w => !string.IsNullOrWhiteSpace(w.NewMtlItemNo)).Select(s => s.NewMtlItemNo).Distinct().ToArray();
                    string[] newMtlItemNames = newBomList.Where(w => !string.IsNullOrWhiteSpace(w.NewMtlItemName)).Select(s => s.NewMtlItemName).Distinct().ToArray();
                    int[] newInventoryIds = newBomList.Where(w => w.NewInventoryId.HasValue && w.NewInventoryId > 0).Select(s => s.NewInventoryId.Value).Distinct().ToArray();

                    #region //取得品號資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT MtlItemNo
                            FROM PDM.MtlItem
                            WHERE MtlItemNo IN @MtlItemNos
                            AND CompanyId = @CompanyId";
                    dynamicParameters.Add("MtlItemNos", newMtlItemNos);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    var resultMtlItemNos = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //取得品號資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT MtlItemName
                            FROM PDM.MtlItem
                            WHERE MtlItemName IN @MtlItemNames
                            AND CompanyId = @CompanyId";
                    dynamicParameters.Add("MtlItemNames", newMtlItemNames);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    var resultMtlItemNames = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //取得庫別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT InventoryId
                            FROM SCM.Inventory
                            WHERE InventoryId IN @InventoryIds
                            AND CompanyId = @CompanyId";
                    dynamicParameters.Add("InventoryIds", newInventoryIds);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    var resultInventoryIds = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //取得品號公司別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT CompanyId, CompanyNo, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var resultCorp = sqlConnection.Query(sql, dynamicParameters);
                    string companyNo = string.Empty;
                    foreach (var corp in resultCorp)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[corp.ErpDb];
                        companyNo = corp.CompanyNo;
                    }
                    #endregion

                    List<dynamic> resultINVMB = new List<dynamic>();
                    List<Erp> resultINVMA = new List<Erp>();
                    string MA037 = string.Empty;
                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //取得ERP品號 INVMB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MB001)) AS MB001
                                FROM INVMB
                                WHERE COMPANY = @CompanyNo 
                                AND RTRIM(LTRIM(MB001)) IN @MB001s";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("MB001s", newMtlItemNos);
                        resultINVMB = erpConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取類別一、二、三、四
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MA002)) MtlItemTypeNo, LTRIM(RTRIM(MA003)) MtlItemTypeName
                                FROM INVMA
                                WHERE 1=1";
                        resultINVMA = erpConnection.Query<Erp>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //ERP共用參數設定檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MA037 FROM CMSMA";
                        var resultCMSMA = erpConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCMSMA)
                        {
                            MA037 = item.MA037;
                        }
                        #endregion
                    }

                    List<string> msg = new List<string>();
                    List<BillOfMaterialsTree> errorList = new List<BillOfMaterialsTree>();

                    #region //品號檢查
                    newBomList.ForEach(item =>
                    {
                        item.Msg = string.Empty; //初始化訊息
                        msg = new List<string>();

                        if (item.IsNewItem && string.IsNullOrWhiteSpace(item.MtlItemNo))
                            msg.Add("請輸入品號");
                        if (!string.IsNullOrWhiteSpace(item.NewMtlItemNo) && newBomList.Any(a => a.MtlItemId != item.MtlItemId && a.NewMtlItemNo == item.NewMtlItemNo))
                            msg.Add("重複使用");

                        var d = resultMtlItemNos.FirstOrDefault(f => f.MtlItemNo == item.NewMtlItemNo);
                        if (d != null) msg.Add("MES已存在該品號");

                        var d1 = resultMtlItemNames.FirstOrDefault(f => f.MtlItemName == item.NewMtlItemName);
                        if (d1 != null && MA037 == "1") msg.Add("MES已存在該品名");

                        var e = resultINVMB.FirstOrDefault(f => f.MB001 == item.NewMtlItemNo);
                        if (e != null) msg.Add("ERP已存在該品號");

                        var e1 = resultINVMB.FirstOrDefault(f => f.MB002 == item.NewMtlItemName);
                        if (e1 != null && MA037 == "1") msg.Add("ERP已存在該品名");

                        if (item.IsNewItem)
                        {
                            //24.08.16 僅能包含半形大寫字母/半形數字/減號
                            if (!Regex.IsMatch(item.NewMtlItemNo, @"^[A-Z0-9-]+$"))
                                msg.Add("僅能包含:半形大寫字母|半形數字|減號");

                            //24.01.17 加入(是否包含小寫字母)判斷
                            if (Regex.IsMatch(item.NewMtlItemNo, @"[a-z]"))
                                msg.Add("須大寫");

                            //24.01.05 加入(有無寬度的空白字元)判斷
                            if (Regex.IsMatch(item.NewMtlItemNo, @"[\s\u0009-\u000D\u0020\u0085\u00A0\u1680\u2000-\u200A\u2028-\u2029\u202F\u205F\u3000\u180E\u200B-\u200D\u2060\uFEFF\u00B7\u237D\u2420\u2422\u2423]"))
                                msg.Add("不可空白");

                            //24.01.05 加入(全形數字 大寫英文 小寫英文)判斷
                            if (Regex.IsMatch(item.NewMtlItemNo, @"[\uFF10-\uFF19\uFF41-\uFF5A\uFF21-\uFF3A]"))
                                msg.Add("須半形文字");

                            //MtlItemNo是否包含繁簡體中文
                            if (Regex.IsMatch(item.NewMtlItemNo, @"\p{IsCJKUnifiedIdeographs}") == true)
                                msg.Add("不可簡體");
                            if (Regex.IsMatch(item.NewMtlItemNo, @"[\u4e00-\u9fa5\u2e80-\u2eff\u2f00-\u2fdf\u3000-\u303f\u31c0-\u31ef\u3200-\u32ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"))
                                msg.Add("不可繁體");
                        }

                        if (!string.IsNullOrWhiteSpace(item.NewTypeOne) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeOne) == null)
                            msg.Add("類別(一)選擇有誤");
                        if (!string.IsNullOrWhiteSpace(item.NewTypeTwo) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeTwo) == null)
                            msg.Add("類別(二)選擇有誤");
                        if (!string.IsNullOrWhiteSpace(item.NewTypeThree) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeThree) == null)
                            msg.Add("類別(三)選擇有誤");
                        if (!string.IsNullOrWhiteSpace(item.NewTypeFour) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeFour) == null)
                            msg.Add("類別(四)選擇有誤");
                        if (item.NewInventoryId > 0 && resultInventoryIds.FirstOrDefault(f => Convert.ToInt32(f.InventoryId) == item.NewInventoryId) == null)
                            msg.Add("庫別選擇有誤");

                        if (msg.Count() > 0)
                        {
                            errorList.Add(new BillOfMaterialsTree { ParentMtlItemId = item.ParentMtlItemId, MtlItemId = item.MtlItemId, Msg = string.Join(",", msg) });
                        }
                    });
                    #endregion

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = errorList.Where(w => w.MtlItemId > 0)
                    });
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddBomStructure -- 品號BOM結構新增 -- Chia Yuan 2024.07
        public string AddBomStructure(string BomStructureJson)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(BomStructureJson)) throw new SystemException("【結構】資料錯誤!");
                if (!BomStructureJson.TryParseJson(out JObject bomStructureJson)) throw new SystemException("【結構】資料錯誤!");
                if (bomStructureJson["data"].Count() <= 0) throw new SystemException("【結構】資料不存在!");

                List<BillOfMaterialsTree> newBomList = JsonConvert.DeserializeObject<List<BillOfMaterialsTree>>(JObject.Parse(BomStructureJson)["data"].ToString()).Where(w => w.IsClone).ToList(); //反序列化
                if (!newBomList.Any()) throw new SystemException("【結構】資料不存在!");

                var confirmNewBomList = newBomList.Where(w => w.IsNewItem && w.IsClone).ToList(); //確定新建的品號
                if (string.IsNullOrWhiteSpace(confirmNewBomList.FirstOrDefault(f => f.ParentMtlItemId < 0)?.NewMtlItemNo)) throw new SystemException("【主件品號】資料錯誤!");

                var errorList1 = confirmNewBomList.Where(w => string.IsNullOrWhiteSpace(w.NewMtlItemNo)).ToList();
                if (errorList1.Any()) throw new SystemException(string.Format("以下來源必須給予新品號<br/>{0}", string.Join("<br/>", errorList1.Select(s => s.MtlItemNo))));

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string[] newMtlItemNos = newBomList.Where(w => !string.IsNullOrWhiteSpace(w.NewMtlItemNo)).Select(s => s.NewMtlItemNo).Distinct().ToArray();
                        string[] newMtlItemNames = newBomList.Where(w => !string.IsNullOrWhiteSpace(w.NewMtlItemName)).Select(s => s.NewMtlItemName).Distinct().ToArray();
                        int[] newInventoryIds = newBomList.Where(w => w.NewInventoryId.HasValue && w.NewInventoryId > 0).Select(s => s.NewInventoryId.Value).Distinct().ToArray();
                        int[] mtlItemIds = newBomList.Select(s => s.MtlItemId.Value).Distinct().ToArray();

                        #region //取得品號資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT MtlItemNo
                                FROM PDM.MtlItem
                                WHERE MtlItemNo IN @MtlItemNos
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MtlItemNos", newMtlItemNos);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var resultMtlItemNos = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得品名資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT MtlItemName
                                FROM PDM.MtlItem
                                WHERE MtlItemName IN @MtlItemNames
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MtlItemNames", newMtlItemNames);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var resultMtlItemNames = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得庫別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId
                                FROM SCM.Inventory
                                WHERE InventoryId IN @InventoryIds
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("InventoryIds", newInventoryIds);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var resultInventoryIds = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得來源品號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId 
                                , CompanyId, ItemTypeId, MtlItemNo, MtlItemName, MtlItemSpec
                                , InventoryUomId, PurchaseUomId, SaleUomId
                                , TypeOne, TypeTwo, TypeThree, TypeFour
                                , InventoryId, RequisitionInventoryId, InventoryManagement, MtlModify, BondedStore, ItemAttribute, ProductionLine, LotManagement
                                , EfficientDays, RetestDays, MeasureType, EfficientDays, ExpirationDate, OverReceiptManagement, OverDeliveryManagement
                                , [Version], MtlItemDesc, MtlItemRemark, TransferStatus, TransferDate, [Status]
                                FROM PDM.MtlItem
                                WHERE MtlItemId IN @MtlItemIds";
                        dynamicParameters.Add("MtlItemIds", mtlItemIds);
                        var sourceMtlItems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得來源Bom
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT BomId, MtlItemId
                                FROM PDM.BillOfMaterials
                                WHERE MtlItemId IN @MtlItemIds";
                        dynamicParameters.Add("MtlItemIds", mtlItemIds);
                        var sourceBoms = sqlConnection.Query<BillOfMaterials>(sql, dynamicParameters).ToList();
                        #endregion

                        var bomIds = sourceBoms.Select(s => s.BomId).Distinct().ToArray();

                        #region //取得來源BomDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT BomId, BomDetailId, MtlItemId
                                FROM PDM.BomDetail
                                WHERE BomId IN @BomIds";
                        dynamicParameters.Add("BomIds", bomIds);
                        var sourceBomDetails = sqlConnection.Query<BomDetail>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得品號公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT CompanyId, CompanyNo, CompanyName, ErpDb
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCorp = sqlConnection.Query(sql, dynamicParameters);
                        string companyNo = string.Empty;
                        foreach (var corp in resultCorp)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[corp.ErpDb];
                            companyNo = corp.CompanyNo;
                        }
                        #endregion

                        #region //撈取單位ID
                        int uomId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId 
                                FROM PDM.UnitOfMeasure
                                WHERE UomNo = 'PCS'
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultUnit = sqlConnection.Query(sql, dynamicParameters);

                        if (resultUnit.Count() <= 0) throw new SystemException("【品號管理】－單位資料錯誤!");

                        foreach (var item in resultUnit)
                        {
                            uomId = item.UomId;
                        }
                        #endregion

                        List<dynamic> resultINVMB = new List<dynamic>();
                        List<Erp> resultINVMA = new List<Erp>();
                        string MA037 = string.Empty;
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得ERP品號 INVMB
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB001)) AS MB001
                                    FROM INVMB
                                    WHERE COMPANY = @CompanyNo 
                                    AND RTRIM(LTRIM(MB001)) IN @MB001s";
                            dynamicParameters.Add("CompanyNo", companyNo);
                            dynamicParameters.Add("MB001s", newMtlItemNos);
                            resultINVMB = erpConnection.Query(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取類別一、二、三、四
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA002)) MtlItemTypeNo, LTRIM(RTRIM(MA003)) MtlItemTypeName
                                    FROM INVMA
                                    WHERE 1=1";
                            resultINVMA = erpConnection.Query<Erp>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //ERP共用參數設定檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MA037 FROM CMSMA";
                            var resultCMSMA = erpConnection.Query(sql, dynamicParameters);
                            foreach (var item in resultCMSMA)
                            {
                                MA037 = item.MA037;
                            }
                            #endregion
                        }

                        List<string> msg = new List<string>();
                        List<BillOfMaterialsTree> errorList = new List<BillOfMaterialsTree>();

                        #region //品號檢查
                        newBomList.ForEach(item =>
                        {
                            item.Msg = string.Empty; //初始化訊息
                            msg = new List<string>();

                            if (item.IsNewItem && string.IsNullOrWhiteSpace(item.MtlItemNo))
                                msg.Add("請輸入品號");
                            if (!string.IsNullOrWhiteSpace(item.NewMtlItemNo) && newBomList.Any(a => a.MtlItemId != item.MtlItemId && a.NewMtlItemNo == item.NewMtlItemNo))
                                msg.Add("重複使用");

                            var d = resultMtlItemNos.FirstOrDefault(f => f.MtlItemNo == item.NewMtlItemNo);
                            if (d != null) msg.Add("MES已存在該品號");

                            var d1 = resultMtlItemNames.FirstOrDefault(f => f.MtlItemName == item.NewMtlItemName);
                            if (d1 != null && MA037 == "1") msg.Add("MES已存在該品名");

                            var e = resultINVMB.FirstOrDefault(f => f.MB001 == item.NewMtlItemNo);
                            if (e != null) msg.Add("ERP已存在該品號");

                            var e1 = resultINVMB.FirstOrDefault(f => f.MB002 == item.NewMtlItemName);
                            if (e1 != null && MA037 == "1") msg.Add("ERP已存在該品名");

                            if (item.IsNewItem)
                            {
                                //24.08.16 僅能包含半形大寫字母/半形數字/減號
                                if (!Regex.IsMatch(item.NewMtlItemNo, @"^[A-Z0-9-]+$"))
                                    msg.Add("僅能包含:半形大寫字母|半形數字|減號");

                                //24.01.17 加入(是否包含小寫字母)判斷
                                if (Regex.IsMatch(item.NewMtlItemNo, @"[a-z]"))
                                    msg.Add("須大寫");

                                //24.01.05 加入(有無寬度的空白字元)判斷
                                if (Regex.IsMatch(item.NewMtlItemNo, @"[\s\u0009-\u000D\u0020\u0085\u00A0\u1680\u2000-\u200A\u2028-\u2029\u202F\u205F\u3000\u180E\u200B-\u200D\u2060\uFEFF\u00B7\u237D\u2420\u2422\u2423]"))
                                    msg.Add("不可空白");

                                //24.01.05 加入(全形數字 大寫英文 小寫英文)判斷
                                if (Regex.IsMatch(item.NewMtlItemNo, @"[\uFF10-\uFF19\uFF41-\uFF5A\uFF21-\uFF3A]"))
                                    msg.Add("須半形文字");

                                //MtlItemNo是否包含繁簡體中文
                                if (Regex.IsMatch(item.NewMtlItemNo, @"\p{IsCJKUnifiedIdeographs}") == true)
                                    msg.Add("不可簡體");
                                if (Regex.IsMatch(item.NewMtlItemNo, @"[\u4e00-\u9fa5\u2e80-\u2eff\u2f00-\u2fdf\u3000-\u303f\u31c0-\u31ef\u3200-\u32ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"))
                                    msg.Add("不可繁體");
                            }

                            if (!string.IsNullOrWhiteSpace(item.NewTypeOne) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeOne) == null)
                                msg.Add("類別(一)選擇有誤");
                            if (!string.IsNullOrWhiteSpace(item.NewTypeTwo) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeTwo) == null)
                                msg.Add("類別(二)選擇有誤");
                            if (!string.IsNullOrWhiteSpace(item.NewTypeThree) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeThree) == null)
                                msg.Add("類別(三)選擇有誤");
                            if (!string.IsNullOrWhiteSpace(item.NewTypeFour) && resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.NewTypeFour) == null)
                                msg.Add("類別(四)選擇有誤");
                            if (item.NewInventoryId > 0 && resultInventoryIds.FirstOrDefault(f => Convert.ToInt32(f.InventoryId) == item.NewInventoryId) == null)
                                msg.Add("庫別選擇有誤");

                            if (msg.Count() > 0)
                            {
                                errorList.Add(new BillOfMaterialsTree { ParentMtlItemId = item.ParentMtlItemId, MtlItemId = item.MtlItemId, Msg = string.Join(",", msg) });
                            }
                        });

                        if (errorList.Any(a => !string.IsNullOrWhiteSpace(a.Msg)))
                        {
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = errorList.Where(w => w.MtlItemId > 0)
                            });
                            return jsonResponse.ToString();
                        }
                        #endregion

                        var mainMtlItem = newBomList.Where(f => f.ParentMtlItemId == -1).ToList();
                        //var firstSourceItem = sourceMtlItems.FirstOrDefault(f => f.MtlItemId == firstNewItem.MtlItemId);

                        Recursion(mainMtlItem);

                        void Recursion(List<BillOfMaterialsTree> bomTree) {
                            bomTree.ForEach(t => {
                                MtlItem sourceMtlItem = sourceMtlItems.FirstOrDefault(f => f.MtlItemId == t.MtlItemId); //取得Bom(來源)
                                BillOfMaterials sourceBom = sourceBoms.FirstOrDefault(f => f.MtlItemId == t.MtlItemId); //取得Bom(來源)
                                BillOfMaterialsTree parentBom = newBomList.FirstOrDefault(f => f.MtlItemId == t.ParentMtlItemId); //取得上層Bom 

                                if (t.IsNewItem) //是新品號
                                {
                                    #region //檢查品號是否已新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1 FROM PDM.MtlItem WHERE MtlItemNo = @MtlItemNo";
                                    dynamicParameters.Add("MtlItemNo", t.NewMtlItemNo);
                                    var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                    #endregion

                                    if (result == null)
                                    {
                                        #region //新品號新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PDM.MtlItem (CompanyId, ItemTypeId, MtlItemNo, MtlItemName, MtlItemSpec
                                                , InventoryUomId, PurchaseUomId, SaleUomId, TypeOne, TypeTwo
                                                , TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryManagement
                                                , MtlModify, BondedStore, ItemAttribute, ProductionLine, LotManagement, EfficientDays, RetestDays,ReplenishmentPolicy, MeasureType, EffectiveDate
                                                , ExpirationDate, OverReceiptManagement, OverDeliveryManagement, Version
                                                , MtlItemDesc, MtlItemRemark, TransferStatus, TransferDate, Status
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.MtlItemId,INSERTED.MtlItemNo,INSERTED.InventoryUomId
                                                VALUES (@CompanyId, @ItemTypeId, @MtlItemNo, @MtlItemName, @MtlItemSpec
                                                , @InventoryUomId, @PurchaseUomId, @SaleUomId, @TypeOne, @TypeTwo
                                                , @TypeThree, @TypeFour, @InventoryId, @RequisitionInventoryId, @InventoryManagement
                                                , @MtlModify, @BondedStore, @ItemAttribute, @ProductionLine, @LotManagement, @EfficientDays, @RetestDays,@ReplenishmentPolicy, @MeasureType, @EffectiveDate
                                                , @ExpirationDate, @OverReceiptManagement, @OverDeliveryManagement, @Version
                                                , @MtlItemDesc, @MtlItemRemark, @TransferStatus, @TransferDate, @Status
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                sourceMtlItem.CompanyId,
                                                ItemTypeId = (string)null,
                                                MtlItemNo = string.IsNullOrWhiteSpace(t.NewMtlItemNo) ? sourceMtlItem.MtlItemNo : t.NewMtlItemNo,
                                                MtlItemName = string.IsNullOrWhiteSpace(t.NewMtlItemName) ? sourceMtlItem.MtlItemName : t.NewMtlItemName,
                                                MtlItemSpec = string.IsNullOrWhiteSpace(t.NewMtlItemSpec) ? sourceMtlItem.MtlItemSpec : t.NewMtlItemSpec,
                                                sourceMtlItem.InventoryUomId,
                                                sourceMtlItem.PurchaseUomId,
                                                sourceMtlItem.SaleUomId,
                                                TypeOne = string.IsNullOrWhiteSpace(t.NewTypeOne) ? sourceMtlItem.TypeOne : t.NewTypeOne,
                                                TypeTwo = string.IsNullOrWhiteSpace(t.NewTypeTwo) ? sourceMtlItem.TypeTwo : t.NewTypeTwo,
                                                TypeThree = string.IsNullOrWhiteSpace(t.NewTypeThree) ? sourceMtlItem.TypeThree : t.NewTypeThree,
                                                TypeFour = string.IsNullOrWhiteSpace(t.NewTypeFour) ? sourceMtlItem.TypeFour : t.NewTypeFour,
                                                InventoryId = t.NewInventoryId > 0 ? t.NewInventoryId : sourceMtlItem.InventoryId,
                                                sourceMtlItem.RequisitionInventoryId,
                                                sourceMtlItem.InventoryManagement,
                                                sourceMtlItem.MtlModify,
                                                sourceMtlItem.BondedStore,
                                                sourceMtlItem.ItemAttribute,
                                                sourceMtlItem.ProductionLine,
                                                sourceMtlItem.LotManagement,
                                                sourceMtlItem.EfficientDays,
                                                sourceMtlItem.RetestDays,
                                                sourceMtlItem.ReplenishmentPolicy,
                                                sourceMtlItem.MeasureType,
                                                sourceMtlItem.EffectiveDate,
                                                sourceMtlItem.ExpirationDate,
                                                sourceMtlItem.OverReceiptManagement,
                                                sourceMtlItem.OverDeliveryManagement,
                                                Version = "0000",
                                                sourceMtlItem.MtlItemDesc,
                                                sourceMtlItem.MtlItemRemark,
                                                TransferStatus = "N",
                                                TransferDate = (int?)null,
                                                Status = "A",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                        newBomList.Where(w => w.MtlItemId == t.MtlItemId).ToList().ForEach(f => { f.NewMtlItemId = insertResult.MtlItemId; }); //取得新品號id
                                        rowsAffected += 1;
                                        #endregion

                                        #region //品號設定新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PDM.MtlItemSetting (MtlItemId, a.FixedLeadTime, a.ChangeLeadTime, a.MinPoQty, a.MultiplePoQty, a.MultipleReqQty
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                VALUES (@MtlItemId, 0, 0, 0, 0, 0
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                insertResult.MtlItemId,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        insertResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                        rowsAffected += 1;
                                        #endregion

                                        #region //Bom主件新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PDM.BillOfMaterials (MtlItemId, UomId, StandardLot, WipPrefix, ModiPrefix, ModiNo, ModiSequence, Version
                                                , Remark, ConfirmStatus, ConfirmUserId, ConfirmDate, TransferStatus, TransferDate, Status
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.BomId
                                                VALUES (@MtlItemId, @UomId, @StandardLot, @WipPrefix, @ModiPrefix, @ModiNo, @ModiSequence, @Version
                                                , @Remark, @ConfirmStatus, @ConfirmUserId, @ConfirmDate, @TransferStatus, @TransferDate, @Status
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                MtlItemId = t.NewMtlItemId,
                                                UomId = uomId,
                                                StandardLot = 1,
                                                WipPrefix = 5101,
                                                ModiPrefix = "",
                                                ModiNo = "",
                                                ModiSequence = "",
                                                Version = "0000",
                                                Remark = "",
                                                ConfirmStatus = 'N',
                                                ConfirmUserId = (int?)null,
                                                ConfirmDate = (int?)null,
                                                TransferStatus = 'N',
                                                TransferDate = (int?)null,
                                                Status = 'Y',
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });

                                        insertResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                        newBomList.Where(w => w.MtlItemId == t.MtlItemId).ToList().ForEach(ff => { ff.NewBomId = insertResult.BomId; }); //取得新Bom id
                                        rowsAffected += 1;
                                        #endregion
                                    }

                                    if (parentBom != null)
                                    {
                                        #region //檢查子項是否已新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1 FROM PDM.BomDetail WHERE BomId = @BomId AND MtlItemId = @MtlItemId";
                                        dynamicParameters.Add("BomId", parentBom.NewBomId);
                                        dynamicParameters.Add("MtlItemId", t.NewMtlItemId);
                                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                        #endregion

                                        if (result == null)
                                        {
                                            int CopyFromBomId = parentBom.BomId;
                                            int CopyFromMtlItemId = t.MtlItemId.Value;

                                            #region //目前序號
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT ISNULL(CAST(MAX(BomSequence) AS INT), 0) + 10 MaxSequence
                                                    FROM PDM.BomDetail
                                                    WHERE BomId = @BomId";
                                            dynamicParameters.Add("BomId", parentBom.NewBomId);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            int maxSequence = 0;
                                            foreach (var item in result)
                                            {
                                                maxSequence = Convert.ToInt32(item.MaxSequence);
                                            }
                                            #endregion

                                            #region //上層加入新子項目
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO PDM.BomDetail (BomId, BomSequence, MtlItemId, UomId, CompositionQuantity, Base, LossRate
                                                    , StandardCostingType, MaterialProperties
                                                    , ReleaseItem, EffectiveDate, ExpirationDate, Remark, SubstitutionRemark, ConfirmStatus, ComponentType
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.BomDetailId
                                                    SELECT @BomId, @BomSequence, @MtlItemId, UomId, CompositionQuantity, Base, LossRate
                                                    , StandardCostingType, MaterialProperties
                                                    , ReleaseItem, EffectiveDate, ExpirationDate, Remark, SubstitutionRemark, ConfirmStatus, ComponentType
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM PDM.BomDetail
                                                    WHERE BomId = @CopyFromBomId
                                                    AND MtlItemId = @CopyFromMtlItemId";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    BomId = parentBom.NewBomId,
                                                    BomSequence = string.Format("{0:0000}", maxSequence), //取4位數
                                                    MtlItemId = t.NewMtlItemId,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    CopyFromBomId,
                                                    CopyFromMtlItemId
                                                });
                                            var insertResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                        }
                                        #endregion
                                    }
                                }
                                else if (parentBom != null && parentBom.NewBomId > 0) //是現有品號
                                {
                                    int CopyFromBomDetailId = sourceBomDetails.FirstOrDefault(f => f.BomId == parentBom.BomId && f.MtlItemId == t.MtlItemId).BomDetailId.Value;

                                    #region //檢查子項是否已新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1 FROM PDM.BomDetail WHERE BomId = @BomId AND MtlItemId = @MtlItemId";
                                    dynamicParameters.Add("BomId", parentBom.NewBomId);
                                    dynamicParameters.Add("MtlItemId", t.MtlItemId);
                                    var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                    #endregion

                                    if (result == null)
                                    {
                                        #region //目前序號
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(CAST(MAX(BomSequence) AS INT), 0) + 10 MaxSequence
                                                FROM PDM.BomDetail
                                                WHERE BomId = @BomId";
                                        dynamicParameters.Add("BomId", parentBom.NewBomId);
                                        result = sqlConnection.Query(sql, dynamicParameters);
                                        int maxSequence = 0;
                                        foreach (var item in result)
                                        {
                                            maxSequence = Convert.ToInt32(item.MaxSequence);
                                        }
                                        #endregion

                                        #region //上層加入現有子項目
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PDM.BomDetail (BomId, BomSequence, MtlItemId, UomId, CompositionQuantity, Base, LossRate
                                                , StandardCostingType, MaterialProperties
                                                , ReleaseItem, EffectiveDate, ExpirationDate, Remark, SubstitutionRemark, ConfirmStatus, ComponentType
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.BomDetailId
                                                SELECT @BomId, @BomSequence, MtlItemId, UomId, CompositionQuantity, Base, LossRate
                                                , StandardCostingType, MaterialProperties
                                                , ReleaseItem, EffectiveDate, ExpirationDate, Remark, SubstitutionRemark, ConfirmStatus, ComponentType
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                FROM PDM.BomDetail
                                                WHERE BomDetailId = @CopyFromBomDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                BomId = parentBom.NewBomId,
                                                BomSequence = string.Format("{0:0000}", maxSequence), //取4位數
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy,
                                                CopyFromBomDetailId
                                            });
                                        var insertResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                        #endregion

                                        //最後一層，通常為原料，可能沒有來源Bom
                                    }
                                }

                                Recursion(newBomList.Where(w => w.ParentBomId == t.BomId && w.MtlItemRoute.IndexOf(t.MtlItemRoute) >= 0).ToList());
                            });
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

        #region //AddMtlItemBatchExcel -- 品號批量新增 -- Chia Yuan 2024.08.27
        public string AddMtlItemBatchExcel(string SpreadsheetData, string ItemTypeSegmentList, bool CustomizeMtlItemNo)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(SpreadsheetData)) throw new SystemException("【Excel表單】資料錯誤!");
                if (!SpreadsheetData.TryParseJson(out JObject spreadsheetData)) throw new SystemException("【Excel表單】資料錯誤!");
                if (spreadsheetData["data"].Count() <= 0) throw new SystemException("【Excel表單】資料不存在!");

                if (!ItemTypeSegmentList.TryParseJson(out JObject itemTypeSegmentList)) throw new SystemException("【品號編碼原則】資料錯誤!");
                List<MtlItemBatch> dataList = JsonConvert.DeserializeObject<List<MtlItemBatch>>(spreadsheetData["data"].ToString()).Where(w => !string.IsNullOrWhiteSpace(w.ItemNo)).ToList(); //反序列化
                if (!dataList.Any()) throw new SystemException("【Excel表單】資料不存在!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string[] newMtlItemNos = dataList.Where(w => !string.IsNullOrWhiteSpace(w.MtlItemNo)).Select(s => s.MtlItemNo).Distinct().ToArray();
                        string[] newMtlItemNames = dataList.Where(w => !string.IsNullOrWhiteSpace(w.MtlItemName)).Select(s => s.MtlItemName).Distinct().ToArray();
                        int[] newInventoryIds = dataList.Where(w => !string.IsNullOrWhiteSpace(w.InventoryId.ToString())).Select(s => s.InventoryId.Value).Distinct().ToArray();

                        #region //取得品號資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT MtlItemNo
                                FROM PDM.MtlItem
                                WHERE MtlItemNo IN @MtlItemNos
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MtlItemNos", newMtlItemNos);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var resultMtlItemNos = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得品名資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT MtlItemName
                                FROM PDM.MtlItem
                                WHERE MtlItemName IN @MtlItemNames
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MtlItemNames", newMtlItemNames);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var resultMtlItemNames = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //判定品號類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT ItemTypeId
                                FROM PDM.ItemType
                                WHERE ItemTypeId IN @ItemTypeIds";
                        dynamicParameters.Add("ItemTypeIds", dataList.Where(w => w.ItemTypeId.HasValue).Select(s => s.ItemTypeId).ToArray());
                        var resultItemTypes = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取品號結構設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT ItemTypeId, SortNumber, SegmentType, SegmentValue, SuffixCode
                                FROM PDM.ItemTypeSegment
                                WHERE ItemTypeId IN @ItemTypeIds
                                ORDER BY SortNumber";
                        dynamicParameters.Add("ItemTypeIds", dataList.Where(w => w.ItemTypeId.HasValue).Select(s => s.ItemTypeId).ToArray());
                        var resultItemTypeSegments = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得編碼節段主檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT SegmentNo, SegmentName
                                FROM PDM.ItemSegment
                                WHERE ItemSegmentId IN @ItemSegmentIds";
                        dynamicParameters.Add("ItemSegmentIds", resultItemTypeSegments.Where(w => w.SegmentType == "SV" && w.SegmentValue != null).Select(s => s.SegmentValue).ToArray());
                        var resultItemSegments = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得庫別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId
                                FROM SCM.Inventory
                                WHERE InventoryId IN @InventoryIds
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("InventoryIds", newInventoryIds);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var resultInventoryIds = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得品號公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT CompanyId, CompanyNo, CompanyName, ErpDb
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCorp = sqlConnection.Query(sql, dynamicParameters);
                        string companyNo = string.Empty;
                        foreach (var corp in resultCorp)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[corp.ErpDb];
                            companyNo = corp.CompanyNo;
                        }
                        #endregion

                        #region //撈取單位ID
                        int uomId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId 
                                FROM PDM.UnitOfMeasure
                                WHERE UomNo = 'PCS'
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultUnit = sqlConnection.Query(sql, dynamicParameters);

                        if (resultUnit.Count() <= 0) throw new SystemException("【品號管理】－單位資料錯誤!");

                        foreach (var item in resultUnit)
                        {
                            uomId = item.UomId;
                        }
                        #endregion

                        List<dynamic> resultINVMB = new List<dynamic>();
                        List<Erp> resultINVMA = new List<Erp>();
                        string MA037 = string.Empty;
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得ERP品號 INVMB
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB001)) AS MB001
                                    FROM INVMB
                                    WHERE COMPANY = @CompanyNo 
                                    AND RTRIM(LTRIM(MB001)) IN @MB001s";
                            dynamicParameters.Add("CompanyNo", companyNo);
                            dynamicParameters.Add("MB001s", newMtlItemNos);
                            resultINVMB = erpConnection.Query(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取類別一、二、三、四
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA002)) MtlItemTypeNo, LTRIM(RTRIM(MA003)) MtlItemTypeName
                                    FROM INVMA
                                    WHERE 1=1";
                            resultINVMA = erpConnection.Query<Erp>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //ERP共用參數設定檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MA037 FROM CMSMA";
                            var resultCMSMA = erpConnection.Query(sql, dynamicParameters);
                            foreach (var item in resultCMSMA)
                            {
                                MA037 = item.MA037;
                            }
                            #endregion
                        }

                        List<string> msg = new List<string>();
                        List<MtlItemBatch> errorList = new List<MtlItemBatch>();

                        #region //品號檢查
                        dataList.ForEach(item =>
                        {
                            msg = new List<string>();

                            var d1 = resultMtlItemNames.FirstOrDefault(f => f.MtlItemName == item.MtlItemName);
                            if (d1 != null && MA037 == "1") msg.Add("MES已存在該品名");

                            var e1 = resultINVMB.FirstOrDefault(f => f.MB002 == item.MtlItemNo);
                            if (e1 != null && MA037 == "1") msg.Add("ERP已存在該品名");

                            if (resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.TypeOne) == null)
                                msg.Add("【類別(一)】選擇有誤");
                            if (resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.TypeTwo) == null)
                                msg.Add("【類別(二)】選擇有誤");
                            if (resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.TypeThree) == null)
                                msg.Add("【類別(三)】選擇有誤");
                            if (resultINVMA.FirstOrDefault(f => f.MtlItemTypeNo == item.TypeFour) == null)
                                msg.Add("【類別(四)】選擇有誤");
                            if (resultInventoryIds.FirstOrDefault(f => f.InventoryId == item.InventoryId) == null)
                                msg.Add("【庫別】選擇有誤");

                            #region //套用編碼原則 品號格式判斷
                            if (!string.IsNullOrWhiteSpace(item.ItemTypeId.ToString()) && CustomizeMtlItemNo)
                            {
                                var t = resultItemTypes.FirstOrDefault(f => f.ItemTypeId == item.ItemTypeId);
                                if (t == null) msg.Add("【品號類別】錯誤");

                                var t2 = resultItemTypeSegments.FirstOrDefault(f => f.ItemTypeId == item.ItemTypeId);
                                if (t2 == null) msg.Add("【品號結構設定】錯誤");

                                var itemTypeSegmentValue = itemTypeSegmentList["data"].FirstOrDefault(f => Convert.ToInt32(f["row"].ToString()) == item.row); //編碼節段資料
                                var itemTypeSegments = resultItemTypeSegments.Where(w => w.ItemTypeId == item.ItemTypeId);

                                #region //取得品號編碼組合 SegmentType <> S
                                List<string> SaveValues = new List<string>();
                                foreach (var segment in itemTypeSegments.Where(w => w.SegmentType != "S").ToList())
                                {
                                    string saveValue = itemTypeSegmentValue["SegmentType" + segment.SegmentType + segment.SortNumber]?.ToString() + (segment.SuffixCode == "Y" ? "-" : string.Empty);
                                    if (SaveValues.FirstOrDefault(f => f == saveValue) == null) SaveValues.Add(saveValue); //取得實際節段編碼值
                                }
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MtlItemId, COUNT(MtlItemId) 'SegmentCount'
                                        --, SegmentType, SegmentSort, SaveValue
                                        FROM PDM.MtlItemSegment
                                        WHERE SegmentType IN @SegmentType
                                        AND SegmentSort IN @SegmentSort
                                        AND SaveValue IN @SaveValue
                                        GROUP BY MtlItemId --, SegmentType, SegmentSort, SaveValue
                                        HAVING(COUNT(MtlItemId) = @SegmentCount)
                                        ORDER BY MtlItemId DESC";
                                dynamicParameters.Add("SegmentType", itemTypeSegments.Where(w => w.SegmentType != "S").Select(s => s.SegmentType).ToArray());
                                dynamicParameters.Add("SegmentSort", itemTypeSegments.Where(w => w.SegmentType != "S").Select(s => s.SortNumber).ToArray());
                                dynamicParameters.Add("SaveValue", SaveValues.ToArray());
                                dynamicParameters.Add("SegmentCount", itemTypeSegments.Where(w => w.SegmentType != "S").Count());
                                var segmentMtlItemIds = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                int maxSequence = 0;
                                string tempMtlItemNo = string.Empty;

                                #region //根據品號結構產生新品號+判斷是否符合設定
                                foreach (var segment in itemTypeSegments) //品號結構解析
                                {
                                    string saveValue = itemTypeSegmentValue["SegmentType" + segment.SegmentType + segment.SortNumber]?.ToString();
                                    string segmentValue = saveValue;

                                    switch (segment.SegmentType)
                                    {
                                        case "C": //自訂義
                                        case "D": //日期類
                                        case "F": //固定值
                                        case "SV": //編碼節段
                                            if (segment.SegmentType == "C")
                                            {
                                                if (string.IsNullOrWhiteSpace(saveValue)) { msg.Add("【自訂節段碼】未編輯!"); continue; }
                                                if (saveValue.Length != Convert.ToInt32(segment.SegmentValue)) msg.Add("【自訂節段碼】須" + segment.SegmentValue + "字元!");
                                            }
                                            if (segment.SegmentType == "SV")
                                            {
                                                if (string.IsNullOrWhiteSpace(saveValue)) { msg.Add("【編碼節段】未選取!"); continue; }
                                            }
                                            saveValue = saveValue + (segment.SuffixCode == "Y" ? "-" : string.Empty);
                                            tempMtlItemNo += saveValue;
                                            break;
                                        case "S": //流水號碼數
                                            if (string.IsNullOrWhiteSpace(item.MtlItemNo)) { msg.Add("【品號】格式錯誤!"); continue; }

                                            char charToCount = '?';
                                            int SequenceNum = item.MtlItemNo.Count(c => c == charToCount); //流水號長度

                                            #region //撈取該組合目前最大值 (非品號編碼原則狀況)
                                            string Qmark = "";
                                            string direction = "One";
                                            for (var i = 0; i < SequenceNum; i++)
                                            {
                                                Qmark += "?";
                                            }
                                            int Noplace = item.MtlItemNo.IndexOf(Qmark);
                                            int NoLong = item.MtlItemNo.Length - SequenceNum;
                                            string[] splitStr = { Qmark }; //自行設定切割字串

                                            string StrBefore = item.MtlItemNo.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[0];
                                            string StrAfter = "";
                                            if (Noplace > 0 || Noplace == NoLong)
                                            {
                                                if (segment.SuffixCode == "Y")
                                                {
                                                    StrAfter = item.MtlItemNo.Split(splitStr, StringSplitOptions.RemoveEmptyEntries)[1];
                                                    direction = "Two";
                                                }
                                            }

                                            switch (direction)
                                            {
                                                case "Two":
                                                    var dataCust2 = dataList.Where(w => !w.MtlItemNo.Contains(charToCount) && w.MtlItemNo.StartsWith(StrBefore) && w.MtlItemNo.EndsWith(StrAfter))
                                                        .Select(s => s.MtlItemNo.Replace(StrBefore, string.Empty).Replace(StrAfter, string.Empty)).Max();
                                                    int.TryParse(dataCust2, out int maxSequenceoCust2);

                                                   dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT MAX(REPLACE(REPLACE(MtlItemNo, @StrBefore, ''), @StrAfter, '')) maxSequence
                                                            FROM PDM.MtlItem
                                                            WHERE MtlItemNo LIKE @StrBefore +'%'
                                                            AND MtlItemNo LIKE + '%' + @StrAfter
                                                            AND MtlItemNo NOT LIKE '%(複製)%'
                                                            AND LEN(MtlItemNo) = @MtlItemLenght";
                                                    dynamicParameters.Add("StrBefore", StrBefore);
                                                    dynamicParameters.Add("StrAfter", StrAfter);
                                                    dynamicParameters.Add("MtlItemLenght", item.MtlItemNo.Length);
                                                    var resultMaxSequence2 = sqlConnection.Query(sql, dynamicParameters);
                                                    foreach (var seq in resultMaxSequence2)
                                                    {
                                                        maxSequence = Convert.ToInt32(seq.maxSequence);
                                                    }
                                                    if (maxSequenceoCust2 > maxSequence) maxSequence = maxSequenceoCust2;
                                                    break;
                                                case "One":
                                                    var dataCust1 = dataList.Where(w => !w.MtlItemNo.Contains(charToCount) && w.MtlItemNo.StartsWith(StrBefore)).Select(s => s.MtlItemNo.Replace(StrBefore, string.Empty)).Max();
                                                    int.TryParse(dataCust1, out int maxSequenceoCust1);

                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT MAX(REPLACE(MtlItemNo, @StrBefore, '')) maxSequence
                                                            FROM PDM.MtlItem
                                                            WHERE MtlItemNo NOT LIKE '%(複製)%'
                                                            AND LEN(MtlItemNo) = @MtlItemLenght";
                                                    dynamicParameters.Add("StrBefore", StrBefore);
                                                    dynamicParameters.Add("MtlItemLenght", item.MtlItemNo.Length);
                                                    if (Noplace == 0)
                                                    {
                                                        sql += " AND MtlItemNo LIKE '%' + @StrBefore ";
                                                    }
                                                    else if (Noplace == NoLong)
                                                    {
                                                        sql += " AND MtlItemNo LIKE @StrBefore + '%'";
                                                    }
                                                    var resultMaxSequence1 = sqlConnection.Query(sql, dynamicParameters);
                                                    foreach (var seq in resultMaxSequence1)
                                                    {
                                                        maxSequence = Convert.ToInt32(seq.maxSequence);
                                                    }
                                                    if (maxSequenceoCust1 > maxSequence) maxSequence = maxSequenceoCust1;
                                                    break;
                                            }
                                            #endregion

                                            #region //撈取該組合目前最大值 (有品號編碼原則狀況)
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT MAX(REPLACE(SaveValue, '-', '')) AS maxSequence
                                                    FROM PDM.MtlItemSegment
                                                    WHERE MtlItemId IN @MtlItemId
                                                    AND SegmentType = @SegmentType
                                                    AND SegmentSort = @SegmentSort";
                                            dynamicParameters.Add("MtlItemId", segmentMtlItemIds.Select(s => s.MtlItemId).ToArray());
                                            dynamicParameters.Add("SegmentType", segment.SegmentType);
                                            dynamicParameters.Add("SegmentSort", segment.SortNumber);
                                            var mtlItemSegment = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);

                                            if (mtlItemSegment != null) //如果符合規格的列表有值 就要判斷有沒有符合相同規格的流水碼,有的話從中找出目前流水號規則的最大值
                                            {
                                                if (maxSequence < Convert.ToInt32(mtlItemSegment.maxSequence))
                                                    maxSequence = Convert.ToInt32(mtlItemSegment.maxSequence);

                                                SequenceNum = Convert.ToInt32(segment.SegmentValue); //流水號長度(從組合表)
                                            }
                                            else  //找不到代表沒有這個組合,所以流水號從1開始
                                            {
                                                if (maxSequence <= 0) maxSequence = 1;
                                                else maxSequence += 1;
                                            }
                                            #endregion

                                            saveValue = (maxSequence + 1).ToString().PadLeft(SequenceNum, '0');
                                            if (saveValue.Length > SequenceNum)
                                            {
                                                msg.Add("目前流水編號已經到達到最大值" + SequenceNum + "字元!");
                                            }
                                            segmentValue = saveValue;
                                            saveValue = saveValue + (segment.SuffixCode == "Y" ? "-" : string.Empty);
                                            tempMtlItemNo += saveValue;
                                            break;
                                    }

                                    //產生品號編碼節段資料
                                    item.MtlItemSegments.Add(new MtlItemSegment { ItemTypeId = item.ItemTypeId.Value, SegmentType = segment.SegmentType, SegmentSort = segment.SortNumber, SegmentValue = segmentValue, SaveValue = saveValue, SuffixCode = segment.SuffixCode });
                                }
                                #endregion

                                item.MtlItemNo = tempMtlItemNo;
                            }
                            #endregion
                            #region //不套用編碼原則 品號格式判斷
                            else
                            {
                                //24.08.16 僅能包含半形大寫字母 / 半形數字 / 減號
                                if (!Regex.IsMatch(item.MtlItemNo, @"^[A-Z0-9-]+$"))
                                    msg.Add("【品號】須半形大寫字母|半形數字|減號");

                                //24.01.05 加入(有無寬度的空白字元)判斷
                                if (Regex.IsMatch(item.MtlItemNo, @"[\s\u0009-\u000D\u0020\u0085\u00A0\u1680\u2000-\u200A\u2028-\u2029\u202F\u205F\u3000\u180E\u200B-\u200D\u2060\uFEFF\u00B7\u237D\u2420\u2422\u2423]"))
                                    msg.Add("【品號】不可空白");

                                //24.01.05 加入(全形數字 大寫英文 小寫英文)判斷
                                if (Regex.IsMatch(item.MtlItemNo, @"[\uFF10-\uFF19\uFF41-\uFF5A\uFF21-\uFF3A]"))
                                    msg.Add("【品號】須半形文字");

                                //MtlItemNo是否包含繁簡體中文
                                if (Regex.IsMatch(item.MtlItemNo, @"\p{IsCJKUnifiedIdeographs}") == true)
                                    msg.Add("【品號】不可簡體");
                                if (Regex.IsMatch(item.MtlItemNo, @"[\u4e00-\u9fa5\u2e80-\u2eff\u2f00-\u2fdf\u3000-\u303f\u31c0-\u31ef\u3200-\u32ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"))
                                    msg.Add("【品號】不可繁體");
                            }
                            #endregion

                            //24.08.16 僅能包含半形大寫字母 / 半形數字 / 減號
                            if (!Regex.IsMatch(item.MtlItemNo, @"^[A-Z0-9-]+$") || string.IsNullOrWhiteSpace(item.MtlItemNo))
                                msg.Add("【品號】格式錯誤!");

                            if (!string.IsNullOrWhiteSpace(item.MtlItemNo) && dataList.Any(a => a.ItemNo != item.ItemNo && a.MtlItemNo == item.MtlItemNo))
                                msg.Add("【品號】重複使用");

                            var d = resultMtlItemNos.FirstOrDefault(f => f.MtlItemNo == item.MtlItemNo);
                            if (d != null) msg.Add("MES已存在該品號");

                            var e = resultINVMB.FirstOrDefault(f => f.MB001 == item.MtlItemNo);
                            if (e != null) msg.Add("ERP已存在該品號");

                            if (msg.Count() > 0)
                            {
                                errorList.Add(new MtlItemBatch { row = item.row, ItemNo = item.ItemNo, MtlItemNo = item.MtlItemNo, Msg = string.Join(",", msg) });
                            }
                        }); //Excel列表迴圈(待處理的新品號) end

                        if (errorList.Any(a => !string.IsNullOrWhiteSpace(a.Msg)))
                        {
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = errorList.Where(w => !string.IsNullOrWhiteSpace(w.ItemNo))
                            });
                            return jsonResponse.ToString();
                        }
                        #endregion

                        dataList.ForEach(item => 
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.MtlItem (CompanyId, ItemTypeId, MtlItemNo, MtlItemName, MtlItemSpec
                                    , InventoryUomId, PurchaseUomId, SaleUomId, TypeOne, TypeTwo
                                    , TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryManagement
                                    , MtlModify, BondedStore, ItemAttribute, ProductionLine, LotManagement, EfficientDays, RetestDays,ReplenishmentPolicy, MeasureType, EffectiveDate
                                    , ExpirationDate, OverReceiptManagement, OverDeliveryManagement, Version
                                    , MtlItemDesc, MtlItemRemark, TransferStatus, TransferDate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MtlItemId,INSERTED.MtlItemNo,INSERTED.InventoryUomId
                                    VALUES (@CompanyId, @ItemTypeId, @MtlItemNo, @MtlItemName, @MtlItemSpec
                                    , @InventoryUomId, @PurchaseUomId, @SaleUomId, @TypeOne, @TypeTwo
                                    , @TypeThree, @TypeFour, @InventoryId, @RequisitionInventoryId, @InventoryManagement
                                    , @MtlModify, @BondedStore, @ItemAttribute, @ProductionLine, @LotManagement, @EfficientDays, @RetestDays,@ReplenishmentPolicy, @MeasureType, @EffectiveDate
                                    , @ExpirationDate, @OverReceiptManagement, @OverDeliveryManagement, @Version
                                    , @MtlItemDesc, @MtlItemRemark, @TransferStatus, @TransferDate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    item.ItemTypeId,
                                    item.MtlItemNo,
                                    item.MtlItemName,
                                    item.MtlItemSpec,
                                    InventoryUomId = item.InventoryUomId > 0 ? item.InventoryUomId : null,
                                    PurchaseUomId = item.PurchaseUomId > 0 ? item.PurchaseUomId : null,
                                    SaleUomId = item.SaleUomId > 0 ? item.SaleUomId : null,
                                    item.TypeOne,
                                    item.TypeTwo,
                                    item.TypeThree,
                                    item.TypeFour,
                                    InventoryId = item.InventoryId > 0 ? item.InventoryId : null,
                                    RequisitionInventoryId = item.RequisitionInventoryId > 0 ? item.RequisitionInventoryId : null,
                                    item.InventoryManagement,
                                    item.MtlModify,
                                    item.BondedStore,
                                    item.ItemAttribute,
                                    ProductionLine = string.Empty,
                                    item.LotManagement,
                                    EfficientDays = item.LotManagement != "N" ? item.EfficientDays : 0,
                                    RetestDays = item.LotManagement != "N" ? item.RetestDays : 0,
                                    item.ReplenishmentPolicy,
                                    item.MeasureType,
                                    EffectiveDate = item.EffectiveDate?.ToString("yyyy-MM-dd"),
                                    ExpirationDate = item.ExpirationDate?.ToString("yyyy-MM-dd"),
                                    item.OverReceiptManagement,
                                    item.OverDeliveryManagement,
                                    Version = "0000",
                                    item.MtlItemDesc,
                                    item.MtlItemRemark,
                                    TransferStatus = "N",
                                    TransferDate = (int?)null,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();

                            foreach (var mtlitem in insertResult)
                            {
                                item.MtlItemSegments.ForEach(f => {
                                    #region //品號類型資料新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PDM.MtlItemSegment (MtlItemId, ItemTypeId, SegmentType
                                          , SegmentSort, SegmentValue, SaveValue, SuffixCode
                                          , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                          OUTPUT INSERTED.MtlItemSegmentId
                                          VALUES (@MtlItemId, @ItemTypeId, @SegmentType
                                          , @SegmentSort, @SegmentValue, @SaveValue, @SuffixCode
                                          , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            mtlitem.MtlItemId,
                                            f.ItemTypeId,
                                            f.SegmentType,
                                            f.SegmentSort,
                                            f.SegmentValue,
                                            f.SaveValue,
                                            f.SuffixCode,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertResult1.Count();
                                    #endregion
                                });

                                if (item.BillOfMaterials == "Y")
                                {
                                    #region //Bom主件新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PDM.BillOfMaterials (MtlItemId, UomId, StandardLot, WipPrefix, ModiPrefix, ModiNo, ModiSequence, Version
                                            , Remark, ConfirmStatus, ConfirmUserId, ConfirmDate, TransferStatus, TransferDate, Status
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.BomId
                                            VALUES (@MtlItemId, @UomId, @StandardLot, @WipPrefix, @ModiPrefix, @ModiNo, @ModiSequence, @Version
                                            , @Remark, @ConfirmStatus, @ConfirmUserId, @ConfirmDate, @TransferStatus, @TransferDate, @Status
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            mtlitem.MtlItemId,
                                            UomId = uomId,
                                            StandardLot = 1,
                                            WipPrefix = 5101,
                                            ModiPrefix = "",
                                            ModiNo = "",
                                            ModiSequence = "",
                                            Version = "0000",
                                            Remark = "",
                                            ConfirmStatus = 'N',
                                            ConfirmUserId = (int?)null,
                                            ConfirmDate = (int?)null,
                                            TransferStatus = 'N',
                                            TransferDate = (int?)null,
                                            Status = 'Y',
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult2 = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertResult2.Count();
                                    #endregion
                                }

                                #region //品號設定新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.MtlItemSetting (MtlItemId, a.FixedLeadTime, a.ChangeLeadTime, a.MinPoQty, a.MultiplePoQty, a.MultipleReqQty
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@MtlItemId, 0, 0, 0, 0, 0
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        mtlitem.MtlItemId,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult3 = sqlConnection.Query(sql, dynamicParameters);
                                rowsAffected += insertResult3.Count();
                                #endregion

                                //if (!string.IsNullOrWhiteSpace(item.ItemTypeId.ToString()))
                                //{
                                //    var itemTypeSegmentValue = itemTypeSegmentList["data"].FirstOrDefault(f => Convert.ToInt32(f["row"].ToString()) == item.row); //編碼節段資料
                                //    var itemTypeSegments = resultItemTypeSegments.Where(w => w.ItemTypeId == item.ItemTypeId);

                                //    foreach (var segment in itemTypeSegments) //品號結構解析
                                //    {
                                //        string saveValue = itemTypeSegmentValue["SegmentType" + segment.SegmentType + segment.SortNumber]?.ToString();
                                //    }
                                //}
                            }
                        });

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

        #region //AddBomDetail -- Bom元件新增
        public string AddBomDetail(int BomId, int MtlItemId, string BomSequence, double CompositionQuantity, double Base
            , double LossRate, string StandardCostingType, string MaterialProperties, string EffectiveDate, string ExpirationDate, string Remark)
        {
            try
            {
                if (MtlItemId <= 0) throw new SystemException("【元件品號】不能為空!");
                if (BomSequence.Length <= 0) throw new SystemException("【序號】不能為空!");
                if (BomSequence.Length > 4) throw new SystemException("【序號】長度錯誤!");
                if (CompositionQuantity < 0) throw new SystemException("【單位用量】不能為空!");
                if (Base < 0) throw new SystemException("【底數】不能為空!");
                if (LossRate < 0) throw new SystemException("【損壞率】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (StandardCostingType != "Y" && StandardCostingType != "N") throw new SystemException("【標準成本計算】不能為空!");
                if (MaterialProperties != "1" && MaterialProperties != "2" && MaterialProperties != "3" && MaterialProperties != "4" && MaterialProperties != "5") throw new SystemException("【材料性質】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string TransferStatus = "";
                        int rowsAffected = 0;

                        #region //判斷元件品號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.BomDetail
                                WHERE MtlItemId = @MtlItemId
                                AND BomId = @BomId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);
                        dynamicParameters.Add("BomId", BomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【元件品號】重複，請重新選擇!");
                        #endregion

                        #region //確認元件品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryUomId
                                FROM PDM.MtlItem a 
                                WHERE MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                        int uomId = -1;
                        foreach (var item in MtlItemResult)
                        {
                            uomId = item.InventoryUomId;
                        }
                        #endregion

                        #region //撈取拋轉狀態
                        sql = @"SELECT TransferStatus
                                FROM PDM.BillOfMaterials
                                WHERE BomId = @BomId";
                        dynamicParameters.Add("BomId", BomId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in result3)
                        {
                            TransferStatus = item.TransferStatus;
                        }
                        #endregion

                        if (TransferStatus != "Y")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.BomDetail (BomId, BomSequence, MtlItemId, UomId, CompositionQuantity, Base, LossRate
                                    , StandardCostingType, MaterialProperties
                                    , ReleaseItem, EffectiveDate, ExpirationDate, Remark, SubstitutionRemark, ConfirmStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.BomDetailId
                                    VALUES (@BomId, @BomSequence, @MtlItemId, @UomId, @CompositionQuantity, @Base, @LossRate
                                    , @StandardCostingType, @MaterialProperties
                                    , @ReleaseItem, @EffectiveDate, @ExpirationDate, @Remark, @SubstitutionRemark, @ConfirmStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BomId,
                                    BomSequence,
                                    MtlItemId,
                                    UomId = uomId,
                                    CompositionQuantity,
                                    Base,
                                    LossRate,
                                    StandardCostingType,
                                    MaterialProperties,
                                    ReleaseItem = "",
                                    EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : null,
                                    ExpirationDate = ExpirationDate.Length > 0 ? ExpirationDate : null,
                                    Remark,
                                    SubstitutionRemark = "",
                                    ConfirmStatus = "N",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
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

        #region //TransferToERP --批量拋轉至ERP
        public string TransferToERP(int MtlItemId)
        {
            try
            {
                if (MtlItemId < 0) throw new SystemException("BOM資料拋轉-【品號】資料錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;

                    List<INVMB> mtlItems = new List<INVMB>();
                    List<INVMO> newMtlItems = new List<INVMO>();
                    List<BOMMC> billOfMaterials = new List<BOMMC>();
                    List<BOMMD> bomDetails = new List<BOMMD>();

                    var resultINVMB = new List<dynamic>();
                    var resultINVMO = new List<dynamic>();
                    var resultBOMMC = new List<dynamic>();
                    var resultBOMMD = new List<dynamic>();

                    string adminUser = "N", userGroup = "";
                    string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss");
                    string companyNo = "", departmentNo = "", userNo = "", userName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (!result.Any()) throw new SystemException("【公司別】資料錯誤!");
                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
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
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "MtlItem");
                        dynamicParameters.Add("DetailCode", "data-transfer");
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (!result.Any()) throw new SystemException("【使用者】資料錯誤!");
                        foreach (var item in result)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //取得品號(BOM)
                        dynamicParameters = new DynamicParameters();
                        sql = GetBomRecursionString(ref dynamicParameters, new List<dynamic> { MtlItemId });
                        sql += @"SELECT a.MtlItemId, a.MtlItemNo MB001, a.MtlItemName MB002, a.MtlItemSpec MB003, b.UomNo MB004
                                , a.TypeOne MB005, a.TypeTwo MB006, a.TypeThree MB007, a.TypeFour MB008
                                , a.MtlItemDesc MB009, d.InventoryNo MB017, a.InventoryManagement MB019
                                , a.BondedStore MB020, a.LotManagement MB022, a.EfficientDays MB023, a.RetestDays MB024, a.ItemAttribute MB025, a.MtlItemRemark MB028
                                , ISNULL(FORMAT(a.EffectiveDate, 'yyyyMMdd'),'') MB030, ISNULL(FORMAT(a.ExpirationDate, 'yyyyMMdd'),'') MB031
                                , a.MeasureType MB043, a.OverReceiptManagement MB044, a.MtlModify MB066, a.MtlItemNo MB080
                                , a.OverDeliveryManagement MB087, e.UomNo MB155, f.UomNo MB156, g.InventoryNo MB157, a.Version MB165
                                , a.TransferStatus, a.ReplenishmentPolicy MB034
	                            FROM @tempData x
	                            INNER JOIN PDM.MtlItem a ON a.MtlItemId = x.MtlItemId
                                LEFT JOIN PDM.UnitOfMeasure b ON b.UomId = a.InventoryUomId
                                LEFT JOIN SCM.Inventory d ON d.InventoryId = a.InventoryId
                                LEFT JOIN PDM.UnitOfMeasure e ON e.UomId = a.PurchaseUomId
                                LEFT JOIN PDM.UnitOfMeasure f ON f.UomId = a.SaleUomId
                                LEFT JOIN SCM.Inventory g ON g.InventoryId = a.RequisitionInventoryId
                                WHERE a.TransferStatus = 'N'
                                ORDER BY x.lvl, x.BomSequence";
                        mtlItems = sqlConnection.Query<INVMB>(sql, dynamicParameters).ToList();
                        if (!mtlItems.Any()) throw new SystemException("查無可拋轉之【品號】!");

                        mtlItems.ForEach(item => {
                            item.CREATOR = userNo;
                            item.CREATE_AP = userNo + "PC";
                            item.COMPANY = companyNo;
                            item.USR_GROUP = userGroup;
                            item.CREATE_DATE = dateNow;
                            item.MODIFIER = "";
                            item.MODI_DATE = "";
                            item.FLAG = "1";
                            item.CREATE_TIME = timeNow;
                            item.CREATE_PRID = "BM";
                            item.MODI_TIME = "";
                            item.MODI_AP = "";
                            item.MODI_PRID = "";
                        });
                        #endregion

                        #region //取得新品號資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemNo MO001, a.MtlItemName MO002, a.MtlItemSpec MO003, b.UomNo MO004
                                , a.TypeOne MO005, a.TypeTwo MO006, a.TypeThree MO007, a.TypeFour MO008
                                , a.MtlItemDesc MO009, d.InventoryNo MO017, a.InventoryManagement MO019
                                , a.BondedStore MO020, a.LotManagement MO022,a.EfficientDays MO023,a.RetestDays MO024, a.ItemAttribute MO025, a.MtlItemRemark MO028
                                , ISNULL(FORMAT(a.EffectiveDate, 'yyyyMMdd'),'') MO030, ISNULL(FORMAT(a.ExpirationDate, 'yyyyMMdd'),'') MO031
                                , a.MeasureType MO043, a.OverReceiptManagement MO044, a.MtlModify MO066, a.MtlItemNo MO080
                                , a.OverDeliveryManagement MO087, e.UomNo MO155, f.UomNo MO156, g.InventoryNo MO163, a.ReplenishmentPolicy MO034
                                FROM PDM.MtlItem a
                                LEFT JOIN PDM.UnitOfMeasure b ON b.UomId = a.InventoryUomId
                                LEFT JOIN SCM.Inventory d ON d.InventoryId = a.InventoryId
                                LEFT JOIN PDM.UnitOfMeasure e ON e.UomId = a.PurchaseUomId
                                LEFT JOIN PDM.UnitOfMeasure f ON f.UomId = a.SaleUomId
                                LEFT JOIN SCM.Inventory g ON g.InventoryId = a.RequisitionInventoryId
                                WHERE a.MtlItemId IN @MtlItemIds";
                        dynamicParameters.Add("MtlItemIds", mtlItems.Select(s => s.MtlItemId).Distinct().ToArray());
                        newMtlItems = sqlConnection.Query<INVMO>(sql, dynamicParameters).ToList();
                        newMtlItems.ForEach(item => {
                            item.CREATOR = userNo;
                            item.CREATE_AP = userNo + "PC";
                            item.COMPANY = companyNo;
                            item.USR_GROUP = userGroup;
                            item.CREATE_DATE = dateNow;
                            item.MODIFIER = "";
                            item.MODI_DATE = "";
                            item.FLAG = "1";
                            item.CREATE_TIME = timeNow;
                            item.CREATE_PRID = "BM";
                            item.MODI_TIME = "";
                            item.MODI_AP = "";
                            item.MODI_PRID = "";
                        });
                        #endregion

                        #region //取得BOM主件資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.MtlItemNo MC001, c.UomNo MC002, a.StandardLot MC004, a.WipPrefix MC005
                                , a.ModiPrefix MC006, a.ModiNo MC007, a.ModiSequence MC008, a.Version MC009
                                , a.Remark MC010, a.ConfirmStatus MC016, a.TransferStatus
                                FROM PDM.BillOfMaterials a
                                INNER JOIN PDM.MtlItem b ON b.MtlItemId = a.MtlItemId
                                LEFT JOIN PDM.UnitOfMeasure c ON a.UomId = c.UomId
                                WHERE a.MtlItemId IN @MtlItemIds";
                        dynamicParameters.Add("MtlItemIds", mtlItems.Select(s => s.MtlItemId).Distinct().ToArray());
                        billOfMaterials = sqlConnection.Query<BOMMC>(sql, dynamicParameters).ToList();
                        billOfMaterials.ForEach(item =>
                        {
                            item.CREATOR = userNo;
                            item.CREATE_AP = userNo + "PC";
                            item.MC018 = userNo;
                            item.COMPANY = companyNo;
                            item.USR_GROUP = userGroup;
                            item.CREATE_DATE = dateNow;
                            item.MODIFIER = "";
                            item.MODI_DATE = "";
                            item.FLAG = "1";
                            item.CREATE_TIME = timeNow;
                            item.CREATE_PRID = "BM";
                            item.MODI_TIME = "";
                            item.MODI_AP = "";
                            item.MODI_PRID = "";
                        });
                        #endregion

                        #region //取得BOM元件資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT c.MtlItemNo MD001, a.BomSequence MD002, d.MtlItemNo MD003, e.UomNo MD004
                                , a.CompositionQuantity MD006, a.Base MD007, a.LossRate MD008
                                , ISNULL(FORMAT(a.EffectiveDate, 'yyyyMMdd'),'') MD011, ISNULL(FORMAT(a.ExpirationDate, 'yyyyMMdd'),'') MD012
                                , a.Remark MD016,a.StandardCostingType MD014,a.MaterialProperties MD017, a.ComponentType UDF01
                                FROM PDM.BomDetail a
                                INNER JOIN PDM.BillOfMaterials b ON b.BomId = a.BomId
                                INNER JOIN PDM.MtlItem c ON c.MtlItemId = b.MtlItemId
                                INNER JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
                                INNER JOIN PDM.UnitOfMeasure e ON e.UomId = a.UomId
                                WHERE b.MtlItemId IN @MtlItemIds";
                        dynamicParameters.Add("MtlItemIds", mtlItems.Select(s => s.MtlItemId).Distinct().ToArray());
                        bomDetails = sqlConnection.Query<BOMMD>(sql, dynamicParameters).ToList();
                        bomDetails.ForEach(item =>
                        {
                            item.CREATOR = userNo;
                            item.CREATE_AP = userNo + "PC";
                            item.COMPANY = companyNo;
                            item.USR_GROUP = userGroup;
                            item.CREATE_DATE = dateNow;
                            item.MODIFIER = "";
                            item.MODI_DATE = "";
                            item.FLAG = "1";
                            item.CREATE_TIME = timeNow;
                            item.CREATE_PRID = "BM";
                            item.MODI_TIME = "";
                            item.MODI_AP = "";
                            item.MODI_PRID = "";
                        });
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得ERP品號 INVMB
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB001)) AS MB001
                                    FROM INVMB
                                    WHERE COMPANY = @CompanyNo 
                                    AND RTRIM(LTRIM(MB001)) IN @MB001s";
                            dynamicParameters.Add("CompanyNo", companyNo);
                            dynamicParameters.Add("MB001s", mtlItems.Select(s => s.MB001).Distinct().ToArray());
                            resultINVMB = erpConnection.Query(sql, dynamicParameters).ToList();
                            #endregion

                            #region //取得ERP新品號 INVMO
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MO001)) AS MO001
                                    FROM INVMO
                                    WHERE COMPANY = @CompanyNo 
                                    AND RTRIM(LTRIM(MO001)) IN @MO001s";
                            dynamicParameters.Add("CompanyNo", companyNo);
                            dynamicParameters.Add("MO001s", newMtlItems.Select(s => s.MO001).Distinct().ToArray());
                            resultINVMO = erpConnection.Query(sql, dynamicParameters).ToList();
                            #endregion

                            #region //取得ERP BOM主件 BOMMC
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MC001)) AS MC001
                                    FROM BOMMC 
                                    WHERE COMPANY = @CompanyNo 
                                    AND RTRIM(LTRIM(MC001)) IN @MC001s";
                            dynamicParameters.Add("CompanyNo", companyNo);
                            dynamicParameters.Add("MC001s", billOfMaterials.Select(s => s.MC001).Distinct().ToArray());
                            resultBOMMC = erpConnection.Query(sql, dynamicParameters).ToList();
                            #endregion

                            #region //判斷BOM元件是否重複 BOMMD
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MD001)) AS MD001, LTRIM(RTRIM(MD003)) AS MD003
                                    FROM BOMMD
                                    WHERE COMPANY = @CompanyNo 
                                    AND RTRIM(LTRIM(MD001)) IN @MD001s 
                                    AND RTRIM(LTRIM(MD003)) IN @MD003s";
                            dynamicParameters.Add("CompanyNo", companyNo);
                            dynamicParameters.Add("MD001s", billOfMaterials.Select(s => s.MC001).Distinct().ToArray());
                            dynamicParameters.Add("MD003s", bomDetails.Select(s => s.MD003).Distinct().ToArray());
                            resultBOMMD = erpConnection.Query(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        var msg = new List<string>();
                        var errorList = new List<string>();

                        #region //品號檢查
                        mtlItems.ForEach(item =>
                        {
                            msg = new List<string>();

                            var e = resultINVMB.FirstOrDefault(f => f.MB001 == item.MB001);
                            if (e != null) msg.Add("ERP已存在此品號");

                            var e2 = resultINVMO.FirstOrDefault(f => f.MO001 == item.MB001);
                            if (e2 != null) msg.Add("ERP已存在此新品號");

                            var e3 = resultBOMMC.FirstOrDefault(f => f.MC001 == item.MB001);
                            if (e3 != null) msg.Add("ERP已存在此主件");

                            var bom = billOfMaterials.FirstOrDefault(f => f.MC001 == item.MB001);
                            if (bom.TransferStatus == "Y" && e3 == null) msg.Add(string.Format(@"MES BOM主件{0}狀態已拋轉，但ERP沒有資料。", item.MB001));

                            var mc001s = billOfMaterials.Select(s => s.MC001).Distinct().ToList();
                            var e4 = resultBOMMD.FirstOrDefault(f => mc001s.Contains(f.MD001) && f.MD003 == item.MB001);
                            if (e4 != null) msg.Add("ERP已存在此元件");

                            //24.08.16 僅能包含半形大寫字母/半形數字/減號
                            if (!Regex.IsMatch(item.MB001, @"^[A-Z0-9-]+$"))
                                msg.Add("僅能包含:半形大寫字母|半形數字|減號");

                            //24.01.17 加入(是否包含小寫字母)判斷
                            if (Regex.IsMatch(item.MB001, @"[a-z]"))
                                msg.Add("須大寫");

                            //24.01.05 加入(有無寬度的空白字元)判斷
                            if (Regex.IsMatch(item.MB001, @"[\s\u0009-\u000D\u0020\u0085\u00A0\u1680\u2000-\u200A\u2028-\u2029\u202F\u205F\u3000\u180E\u200B-\u200D\u2060\uFEFF\u00B7\u237D\u2420\u2422\u2423]"))
                                msg.Add("不可空白");

                            //24.01.05 加入(全形數字 大寫英文 小寫英文)判斷
                            if (Regex.IsMatch(item.MB001, @"[\uFF10-\uFF19\uFF41-\uFF5A\uFF21-\uFF3A]"))
                                msg.Add("須半形文字");

                            //MtlItemNo是否包含繁簡體中文
                            if (Regex.IsMatch(item.MB001, @"\p{IsCJKUnifiedIdeographs}") == true)
                                msg.Add("不可簡體");
                            if (Regex.IsMatch(item.MB001, @"[\u4e00-\u9fa5\u2e80-\u2eff\u2f00-\u2fdf\u3000-\u303f\u31c0-\u31ef\u3200-\u32ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"))
                                msg.Add("不可繁體");

                            if (msg.Any())
                            {
                                errorList.Add(string.Format("{0}:{1}", item.MB001, string.Join(",", msg)));
                            }
                        });

                        if (errorList.Any())
                        {
                            throw new SystemException("以下品號有誤:<br/>" + string.Join("<br/>", errorList));
                        }
                        #endregion
                    }

                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultERP = erpConnection.Query(sql, dynamicParameters);
                        if (!resultERP.Any()) throw new SystemException("【ERP使用者】不存在!");
                        foreach (var item in resultERP)
                        {
                            adminUser = item.MF005; //超級使用者
                            userGroup = item.MF004; //使用者群組
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限 (停用) Function=INVI29, Auth=Y_______Y___
                        //if (adminUser != "Y")
                        //{
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"SELECT TOP 1 MG006
                        //            FROM ADMMG
                        //            WHERE COMPANY = @CompanyNo
                        //            AND MG001 = @UserNo
                        //            AND MG002 = @Function
                        //            AND MG004 = 'Y'
                        //            AND MG006 LIKE @Auth";
                        //    dynamicParameters.Add("CompanyNo", companyNo);
                        //    dynamicParameters.Add("UserNo", userNo);
                        //    dynamicParameters.Add("Function", "INVI29");
                        //    dynamicParameters.Add("Auth", "Y_______Y___"); //修改/查詢/新增
                        //    var resultUserAuthExist = erpConnection.Query(sql, dynamicParameters);
                        //    if (!resultUserAuthExist.Any()) throw new SystemException("品號拋轉-【ERP使用者】無權限!");
                        //}
                        #endregion

                        #region //判斷ERP使用者是否有權限 (停用) Function=BOMI02, Auth=Y_______Y___
                        //if (adminUser != "Y")
                        //{
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"SELECT TOP 1 MG006
                        //            FROM ADMMG
                        //            WHERE COMPANY = @CompanyNo
                        //            AND MG001 = @UserNo
                        //            AND MG002 = @Function
                        //            AND MG004 = 'Y'
                        //            AND MG006 LIKE @Auth";
                        //    dynamicParameters.Add("CompanyNo", companyNo);
                        //    dynamicParameters.Add("UserNo", userNo);
                        //    dynamicParameters.Add("Function", "BOMI02");
                        //    dynamicParameters.Add("Auth", "Y_______Y___"); //修改/查詢/新增
                        //    var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                        //    if (!resultUserAuthExist.Any()) throw new SystemException("BOM拋轉-【ERP使用者】無權限!");
                        //}
                        #endregion

                        var mtlItemIds = mtlItems.Select(s => s.MtlItemId).Distinct().ToList();

                        mtlItemIds.ForEach(mtlItemId =>
                        {
                            var x = mtlItems.FirstOrDefault(f => f.MtlItemId == mtlItemId);

                            switch (x.TransferStatus)
                            {
                                case "N":
                                    #region //品號新增 INVMB from mtlItems
                                    x.MB010 = ""; //標準途程品號
                                    x.MB011 = ""; //標準途程代號
                                    x.MB012 = ""; //文管代號
                                    x.MB013 = ""; //條碼編號
                                    x.MB014 = 0; //單位淨重
                                    x.MB015 = "Kg"; //重量單位
                                    x.MB016 = ""; //外包裝單位
                                    x.MB018 = ""; //計劃人員
                                    x.MB021 = ""; //循環盤點碼
                                    x.MB026 = "99"; //低階碼
                                    x.MB027 = ""; //ABC等級
                                    x.MB029 = ""; //產品圖號
                                    x.MB032 = ""; //主供應商
                                    x.MB033 = "N"; //MPS件
                                    x.MB035 = "2"; //補貨週期
                                    x.MB036 = 0; //固定前置天數
                                    x.MB037 = 0; //變動前置天數
                                    x.MB038 = 0; //批量
                                    x.MB039 = 0; //最低補量
                                    x.MB040 = 0; //補貨倍量
                                    x.MB041 = 0; //領用倍量
                                    x.MB042 = "1"; //領料碼
                                    x.MB045 = 0; //超收率%
                                    x.MB046 = 0; //標準進價
                                    x.MB047 = 0; //標準售價
                                    x.MB048 = ""; //最近進價幣別-原幣別
                                    x.MB049 = 0; //最近進價-原幣單價
                                    x.MB050 = 0; //最近進價-本幣單價
                                    x.MB051 = 0; //零售價
                                    x.MB052 = "N"; //零售價含稅
                                    x.MB053 = 0; //售價定價一
                                    x.MB054 = 0; //售價定價二
                                    x.MB055 = 0; //售價定價三
                                    x.MB056 = 0; //售價定價四
                                    x.MB057 = 0; //單位標準材料成本
                                    x.MB058 = 0; //單位標準人工成本
                                    x.MB059 = 0; //單位標準製造費用
                                    x.MB060 = 0; //單位標準加工費用
                                    x.MB061 = 0; //本階人工
                                    x.MB062 = 0; //本階製費
                                    x.MB063 = 0; //本階加工
                                    x.MB064 = 0; //庫存數量
                                    x.MB065 = 0; //庫存金額
                                    x.MB067 = ""; //採購人員
                                    x.MB068 = ""; //生產線別
                                    x.MB069 = 0; //售價定價五
                                    x.MB070 = 0; //售價定價六
                                    x.MB071 = 0; //外包裝材積
                                    x.MB072 = ""; //小單位
                                    x.MB073 = 0; //外包裝含商品數
                                    x.MB074 = 0; //外包裝淨重
                                    x.MB075 = 0; //外包裝毛重
                                    x.MB076 = 0; //檢驗天數
                                    x.MB077 = ""; //品管類別
                                    x.MB078 = 0; //MRP生產允許交期提前天數
                                    x.MB079 = 0; //MRP採購允許交期提前天數
                                    x.MB081 = ""; //SIZE
                                    x.MB082 = 0; //關稅率
                                    x.MB083 = "N"; //進價管制
                                    x.MB084 = 0; //單價上限率
                                    x.MB085 = "N"; //售價管制
                                    x.MB086 = 0; //單價下限率
                                    x.MB088 = 0; //超交率
                                    x.MB089 = 0; //庫存包裝數量
                                    x.MB090 = ""; //包裝單位
                                    x.MB091 = "N"; //定重
                                    x.MB092 = "N"; //產品序號管理
                                    x.MB093 = 0; //長(CM)
                                    x.MB094 = 0; //寬(CM)
                                    x.MB095 = 0; //高(CM)
                                    x.MB096 = 1; //工時底數
                                    x.MB097 = 0; //業務底價
                                    x.MB098 = "N"; //業務底價含稅
                                    x.MB099 = 0; //貨物稅率
                                    x.MB100 = "N"; //標準進價含稅
                                    x.MB101 = "N"; //標準售價含稅
                                    x.MB102 = "N"; //最近進價-單價含稅(原/本幣)
                                    x.MB103 = "N"; //售價定價一含稅
                                    x.MB104 = "N"; //售價定價二含稅
                                    x.MB105 = "N"; //售價定價三含稅
                                    x.MB106 = "N"; //售價定價四含稅
                                    x.MB107 = "N"; //售價定價五含稅
                                    x.MB108 = "N"; //售價定價六含稅
                                    x.MB109 = "N"; //新品號
                                    x.MB110 = "N"; //料件承認碼
                                    x.MB111 = 0; //轉撥倍量
                                    x.MB112 = ""; //MRP保留欄位
                                    x.MB113 = ""; //APS預留欄位
                                    x.MB114 = ""; //APS預留欄位
                                    x.MB115 = ""; //APS預留欄位
                                    x.MB116 = ""; //APS預留欄位
                                    x.MB117 = ""; //APS預留欄位
                                    x.MB118 = ""; //APS預留欄位
                                    x.MB119 = 0; //APS預留欄位
                                    x.MB120 = 0; //APS預留欄位
                                    x.MB121 = "N"; //控制編碼原則
                                    x.MB122 = ""; //序號前置碼
                                    x.MB123 = ""; //序號流水號碼數
                                    x.MB124 = ""; //序號編碼原則
                                    x.MB125 = ""; //已用生管序號
                                    x.MB126 = ""; //已用商品序號
                                    x.MB127 = ""; //來源
                                    x.MB128 = ""; //預留欄位
                                    x.MB129 = ""; //預留欄位
                                    x.MB130 = ""; //預留欄位
                                    x.MB131 = ""; //電子發票須上傳產品追溯串接碼
                                    x.MB132 = ""; //預留欄位
                                    x.MB133 = ""; //屬性代碼一
                                    x.MB134 = ""; //屬性代碼二
                                    x.MB135 = ""; //屬性代碼三
                                    x.MB136 = ""; //屬性代碼四
                                    x.MB137 = ""; //屬性代碼五
                                    x.MB138 = ""; //屬性代碼六
                                    x.MB139 = ""; //屬性代碼七
                                    x.MB140 = ""; //屬性代碼八
                                    x.MB141 = ""; //屬性代碼九
                                    x.MB142 = ""; //屬性組代碼
                                    x.MB143 = ""; //屬性內容一
                                    x.MB144 = ""; //屬性內容二
                                    x.MB145 = ""; //屬性內容三
                                    x.MB146 = ""; //屬性內容四
                                    x.MB147 = ""; //屬性內容五
                                    x.MB148 = ""; //屬性內容六
                                    x.MB149 = ""; //屬性內容七
                                    x.MB150 = ""; //屬性內容八
                                    x.MB151 = ""; //屬性內容九
                                    x.MB152 = ""; //屬性代碼十
                                    x.MB153 = ""; //屬性內容十
                                    x.MB154 = ""; //圖號版次
                                    x.MB158 = ""; //新品號核准日期
                                    x.MB159 = 0; //預留欄位
                                    x.MB160 = 0; //預留欄位
                                    x.MB161 = 0; //預留欄位
                                    x.MB162 = ""; //預留欄位
                                    x.MB163 = ""; //稅則
                                    x.MB164 = ""; //產品追溯系統串接碼
                                    x.MB166 = 0; //贈品率
                                    x.MB167 = 0; //預留欄位
                                    x.MB168 = ""; //預留欄位
                                    x.MB169 = ""; //DATECODE管理
                                    x.MB170 = ""; //Bin管理
                                    x.MB171 = 0; //預留欄位
                                    x.MB172 = 0; //預留欄位
                                    x.MB173 = 0; //預留欄位
                                    x.MB174 = ""; //預留欄位
                                    x.MB175 = ""; //預留欄位
                                    x.MB176 = ""; //預留欄位
                                    x.MB177 = ""; //預留欄位
                                    x.MB178 = ""; //預留欄位
                                    x.MB179 = ""; //預留欄位
                                    x.MB180 = 0; //APS固定工時
                                    x.MB181 = 0; //APS變動工時
                                    x.MB182 = 0; //批次加工量
                                    x.MB183 = ""; //預留欄位
                                    x.MB184 = 0; //基準數量
                                    x.MB185 = ""; //資源群組
                                    x.MB186 = ""; //資源群組名稱
                                    x.MB187 = ""; //機台代號
                                    x.MB188 = ""; //機台名稱
                                    x.MB189 = ""; //指定資源
                                    x.MB190 = ""; //關鍵料號
                                    x.MB191 = 0; //營業稅率
                                    x.MB192 = 0; //保固佔售價比率
                                    x.MB193 = 0; //保固期數(月數)
                                    x.MB194 = "0"; //交易設限碼
                                    x.MB195 = ""; //建立作業修改日期
                                    x.MB196 = ""; //建立作業修改者
                                    x.MB197 = ""; //MSIC_Code
                                    x.MB198 = ""; //預留欄位
                                    x.MB199 = ""; //預留欄位
                                    x.MB500 = ""; //產品系列ID
                                    x.MB501 = 0; //GROSS_DIE
                                    x.MB502 = ""; //說明1
                                    x.MB503 = ""; //說明2
                                    x.MB504 = ""; //光罩層次
                                    x.MB505 = ""; //進項稅別碼
                                    x.MB506 = ""; //銷項稅別碼
                                    x.MB507 = 0; //預留欄位
                                    x.MB508 = 0; //存貨週轉目標次數
                                    x.MB509 = 0; //預留欄位
                                    x.MB510 = ""; //預留欄位
                                    x.MB511 = ""; //預留欄位
                                    x.MB512 = ""; //預留欄位
                                    x.MB513 = ""; //預留欄位
                                    x.MB514 = ""; //預留欄位
                                    x.MB515 = ""; //開帳批號
                                    x.MB516 = ""; //批號開帳日期
                                    x.MB517 = ""; //內箱標籤格式
                                    x.MB518 = ""; //中箱標籤格式
                                    x.MB519 = ""; //外箱標籤格式
                                    x.MB520 = ""; //內箱選擇列印
                                    x.MB521 = ""; //中箱選擇列印
                                    x.MB522 = ""; //外箱選擇列印
                                    x.MB545 = 0; //標準產出良率
                                    x.MB550 = ""; //產品屬性
                                    x.MB551 = ""; //Wafer型號
                                    x.MB552 = ""; //性質
                                    x.MB553 = ""; //版本
                                    x.MB554 = ""; //作業代號
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
                                    #endregion
                                    #region //INSERT INTO INVMB
                                    sql = @"INSERT INTO INVMB (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , MB001, MB002, MB003, MB004, MB005, MB006, MB007, MB008, MB009, MB010
                                            , MB011, MB012, MB013, MB014, MB015, MB016, MB017, MB018, MB019, MB020
                                            , MB021, MB022, MB023, MB024, MB025, MB026, MB027, MB028, MB029, MB030
                                            , MB031, MB032, MB033, MB034, MB035, MB036, MB037, MB038, MB039, MB040
                                            , MB041, MB042, MB043, MB044, MB045, MB046, MB047, MB048, MB049, MB050
                                            , MB051, MB052, MB053, MB054, MB055, MB056, MB057, MB058, MB059, MB060
                                            , MB061, MB062, MB063, MB064, MB065, MB066, MB067, MB068, MB069, MB070
                                            , MB071, MB072, MB073, MB074, MB075, MB076, MB077, MB078, MB079, MB080
                                            , MB081, MB082, MB083, MB084, MB085, MB086, MB087, MB088, MB089, MB090
                                            , MB091, MB092, MB093, MB094, MB095, MB096, MB097, MB098, MB099, MB100
                                            , MB101, MB102, MB103, MB104, MB105, MB106, MB107, MB108, MB109, MB110
                                            , MB111, MB112, MB113, MB114, MB115, MB116, MB117, MB118, MB119, MB120
                                            , MB121, MB122, MB123, MB124, MB125, MB126, MB127, MB128, MB129, MB130
                                            , MB131, MB132, MB133, MB134, MB135, MB136, MB137, MB138, MB139, MB140
                                            , MB141, MB142, MB143, MB144, MB145, MB146, MB147, MB148, MB149, MB150
                                            , MB151, MB152, MB153, MB154, MB155, MB156, MB157, MB158, MB159, MB160
                                            , MB161, MB162, MB163, MB164, MB165, MB166, MB167, MB168, MB169, MB170
                                            , MB171, MB172, MB173, MB174, MB175, MB176, MB177, MB178, MB179, MB180
                                            , MB181, MB182, MB183, MB184, MB185, MB186, MB187, MB188, MB189, MB190
                                            , MB191, MB192, MB193, MB194, MB195, MB196, MB197, MB198, MB199, MB500
                                            , MB501, MB502, MB503, MB504, MB505, MB506, MB507, MB508, MB509, MB510
                                            , MB511, MB512, MB513, MB514, MB515, MB516, MB517, MB518, MB519, MB520
                                            , MB521, MB522, MB545, MB550, MB551, MB552, MB553, MB554
                                            , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @MB001, @MB002, @MB003, @MB004, @MB005, @MB006, @MB007, @MB008, @MB009, @MB010
                                            , @MB011, @MB012, @MB013, @MB014, @MB015, @MB016, @MB017, @MB018, @MB019, @MB020
                                            , @MB021, @MB022, @MB023, @MB024, @MB025, @MB026, @MB027, @MB028, @MB029, @MB030
                                            , @MB031, @MB032, @MB033, @MB034, @MB035, @MB036, @MB037, @MB038, @MB039, @MB040
                                            , @MB041, @MB042, @MB043, @MB044, @MB045, @MB046, @MB047, @MB048, @MB049, @MB050
                                            , @MB051, @MB052, @MB053, @MB054, @MB055, @MB056, @MB057, @MB058, @MB059, @MB060
                                            , @MB061, @MB062, @MB063, @MB064, @MB065, @MB066, @MB067, @MB068, @MB069, @MB070
                                            , @MB071, @MB072, @MB073, @MB074, @MB075, @MB076, @MB077, @MB078, @MB079, @MB080
                                            , @MB081, @MB082, @MB083, @MB084, @MB085, @MB086, @MB087, @MB088, @MB089, @MB090
                                            , @MB091, @MB092, @MB093, @MB094, @MB095, @MB096, @MB097, @MB098, @MB099, @MB100
                                            , @MB101, @MB102, @MB103, @MB104, @MB105, @MB106, @MB107, @MB108, @MB109, @MB110
                                            , @MB111, @MB112, @MB113, @MB114, @MB115, @MB116, @MB117, @MB118, @MB119, @MB120
                                            , @MB121, @MB122, @MB123, @MB124, @MB125, @MB126, @MB127, @MB128, @MB129, @MB130
                                            , @MB131, @MB132, @MB133, @MB134, @MB135, @MB136, @MB137, @MB138, @MB139, @MB140
                                            , @MB141, @MB142, @MB143, @MB144, @MB145, @MB146, @MB147, @MB148, @MB149, @MB150
                                            , @MB151, @MB152, @MB153, @MB154, @MB155, @MB156, @MB157, @MB158, @MB159, @MB160
                                            , @MB161, @MB162, @MB163, @MB164, @MB165, @MB166, @MB167, @MB168, @MB169, @MB170
                                            , @MB171, @MB172, @MB173, @MB174, @MB175, @MB176, @MB177, @MB178, @MB179, @MB180
                                            , @MB181, @MB182, @MB183, @MB184, @MB185, @MB186, @MB187, @MB188, @MB189, @MB190
                                            , @MB191, @MB192, @MB193, @MB194, @MB195, @MB196, @MB197, @MB198, @MB199, @MB500
                                            , @MB501, @MB502, @MB503, @MB504, @MB505, @MB506, @MB507, @MB508, @MB509, @MB510
                                            , @MB511, @MB512, @MB513, @MB514, @MB515, @MB516, @MB517, @MB518, @MB519, @MB520
                                            , @MB521, @MB522, @MB545, @MB550, @MB551, @MB552, @MB553, @MB554
                                            , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                    rowsAffected += erpConnection.Execute(sql, x);
                                    #endregion

                                    #region //新品號新增 INVMO from newMtlItems
                                    var x2 = newMtlItems.FirstOrDefault(f => f.MO001 == x.MB001);
                                    x2.MO010 = ""; //標準途程品號
                                    x2.MO011 = ""; //標準途程代號
                                    x2.MO012 = ""; //文管代號
                                    x2.MO013 = ""; //條碼編號
                                    x2.MO014 = 0; //單位淨重
                                    x2.MO015 = "Kg"; //重量單位
                                    x2.MO016 = ""; //外包裝單位
                                    x2.MO018 = ""; //計劃人員
                                    x2.MO021 = ""; //循環盤點碼
                                    x2.MO023 = 0; //有效天數
                                    x2.MO024 = 0; //複檢天數
                                    x2.MO026 = ""; //低階碼
                                    x2.MO027 = ""; //ABC等級
                                    x2.MO029 = ""; //產品圖號
                                    x2.MO032 = ""; //主供應商
                                    x2.MO033 = "N"; //MPS件
                                    x2.MO035 = "2"; //補貨週期
                                    x2.MO036 = 0; //固定前置天數
                                    x2.MO037 = 0; //變動前置天數
                                    x2.MO038 = 0; //批量
                                    x2.MO039 = 0; //最低補量
                                    x2.MO040 = 0; //補貨倍量
                                    x2.MO041 = 0; //領用倍量
                                    x2.MO042 = "1"; //領料碼
                                    x2.MO045 = 0; //超收率%
                                    x2.MO046 = 0; //標準進價
                                    x2.MO047 = 0; //標準售價
                                    x2.MO048 = ""; //最近進價幣別-原幣別
                                    x2.MO049 = 0; //最近進價-原幣單價
                                    x2.MO050 = 0; //最近進價-本幣單價
                                    x2.MO051 = 0; //零售價
                                    x2.MO052 = "N"; //零售價含稅
                                    x2.MO053 = 0; //售價定價一
                                    x2.MO054 = 0; //售價定價二
                                    x2.MO055 = 0; //售價定價三
                                    x2.MO056 = 0; //售價定價四
                                    x2.MO057 = 0; //單位標準材料成本
                                    x2.MO058 = 0; //單位標準人工成本
                                    x2.MO059 = 0; //單位標準製造費用
                                    x2.MO060 = 0; //單位標準加工費用
                                    x2.MO061 = 0; //本階人工
                                    x2.MO062 = 0; //本階製費
                                    x2.MO063 = 0; //本階加工
                                    x2.MO064 = 0; //庫存數量
                                    x2.MO065 = 0; //庫存金額
                                    x2.MO067 = ""; //採購人員
                                    x2.MO068 = ""; //生產線別
                                    x2.MO069 = 0; //售價定價五
                                    x2.MO070 = 0; //售價定價六
                                    x2.MO071 = 0; //外包裝材積
                                    x2.MO072 = ""; //小單位
                                    x2.MO073 = 0; //外包裝含商品數
                                    x2.MO074 = 0; //外包裝淨重
                                    x2.MO075 = 0; //外包裝毛重
                                    x2.MO076 = 0; //檢驗天數
                                    x2.MO077 = ""; //品管類別
                                    x2.MO078 = 0; //MRP生產允許交期提前天數
                                    x2.MO079 = 0; //MRP採購允許交期提前天數
                                    x2.MO081 = ""; //SIZE
                                    x2.MO082 = 0; //關稅率
                                    x2.MO083 = "N"; //進價管制
                                    x2.MO084 = 0; //單價上限率
                                    x2.MO085 = "N"; //售價管制
                                    x2.MO086 = 0; //單價下限率
                                    x2.MO088 = 0; //超交率
                                    x2.MO089 = 0; //庫存包裝數量
                                    x2.MO090 = ""; //包裝單位
                                    x2.MO091 = "N"; //定重
                                    x2.MO092 = "N"; //產品序號管理
                                    x2.MO093 = 0; //長(CM)
                                    x2.MO094 = 0; //寬(CM)
                                    x2.MO095 = 0; //高(CM)
                                    x2.MO096 = 0; //工時底數
                                    x2.MO097 = 0; //業務底價
                                    x2.MO098 = "N"; //業務底價含稅
                                    x2.MO099 = 0; //貨物稅率
                                    x2.MO100 = "N"; //標準進價含稅
                                    x2.MO101 = "N"; //標準售價含稅
                                    x2.MO102 = "N"; //最近進價-單價含稅(原/本幣)
                                    x2.MO103 = "N"; //售價定價一含稅
                                    x2.MO104 = "N"; //售價定價二含稅
                                    x2.MO105 = "N"; //售價定價三含稅
                                    x2.MO106 = "N"; //售價定價四含稅
                                    x2.MO107 = "N"; //售價定價五含稅
                                    x2.MO108 = "N"; //售價定價六含稅
                                    x2.MO109 = "N"; //新品號
                                    x2.MO110 = "N"; //料件承認碼
                                    x2.MO111 = 0; //轉撥倍量
                                    x2.MO112 = ""; //MRP保留欄位
                                    x2.MO113 = ""; //APS預留欄位
                                    x2.MO114 = ""; //APS預留欄位
                                    x2.MO115 = ""; //APS預留欄位
                                    x2.MO116 = ""; //APS預留欄位
                                    x2.MO117 = ""; //APS預留欄位
                                    x2.MO118 = ""; //APS預留欄位
                                    x2.MO119 = 0; //APS預留欄位
                                    x2.MO120 = 0; //APS預留欄位
                                    x2.MO121 = "N"; //控制編碼原則
                                    x2.MO122 = ""; //序號前置碼
                                    x2.MO123 = "0"; //序號流水號碼數
                                    x2.MO124 = ""; //序號編碼原則
                                    x2.MO125 = ""; //已用生管序號
                                    x2.MO126 = ""; //已用商品序號
                                    x2.MO127 = ""; //預留欄位
                                    x2.MO128 = ""; //預留欄位
                                    x2.MO129 = ""; //預留欄位
                                    x2.MO130 = ""; //預留欄位
                                    x2.MO131 = ""; //預留欄位
                                    x2.MO132 = ""; //預留欄位
                                    x2.MO133 = ""; //屬性代碼一
                                    x2.MO134 = ""; //屬性代碼二
                                    x2.MO135 = ""; //屬性代碼三
                                    x2.MO136 = ""; //屬性代碼四
                                    x2.MO137 = ""; //屬性代碼五
                                    x2.MO138 = ""; //屬性代碼六
                                    x2.MO139 = ""; //屬性代碼七
                                    x2.MO140 = ""; //屬性代碼八
                                    x2.MO141 = ""; //屬性代碼九
                                    x2.MO142 = ""; //屬性組代碼
                                    x2.MO143 = ""; //屬性內容一
                                    x2.MO144 = ""; //屬性內容二
                                    x2.MO145 = ""; //屬性內容三
                                    x2.MO146 = ""; //屬性內容四
                                    x2.MO147 = ""; //屬性內容五
                                    x2.MO148 = ""; //屬性內容六
                                    x2.MO149 = ""; //屬性內容七
                                    x2.MO150 = ""; //屬性內容八
                                    x2.MO151 = ""; //屬性內容九
                                    x2.MO152 = ""; //屬性代碼十
                                    x2.MO153 = ""; //屬性內容十
                                    x2.MO154 = ""; //圖號版次
                                    x2.MO157 = "Y"; //確認碼
                                    x2.MO158 = dateNow; //確認日期
                                    x2.MO159 = userNo; //確認者
                                    x2.MO160 = "3"; //簽核狀態碼
                                    x2.MO161 = x2.MO001; //品號
                                    x2.MO162 = ""; //申請日期
                                    x2.MO164 = 0; //預留欄位
                                    x2.MO165 = 0; //預留欄位
                                    x2.MO166 = ""; //預留欄位
                                    x2.MO167 = ""; //稅則
                                    x2.MO168 = ""; //預留欄位
                                    x2.MO169 = 0; //贈品率
                                    x2.MO170 = 0; //預留欄位
                                    x2.MO171 = ""; //預留欄位
                                    x2.MO172 = 0; //預留欄位
                                    x2.MO173 = 0; //預留欄位
                                    x2.MO174 = 0; //預留欄位
                                    x2.MO175 = ""; //預留欄位
                                    x2.MO176 = ""; //預留欄位
                                    x2.MO177 = ""; //預留欄位
                                    x2.MO178 = ""; //預留欄位
                                    x2.MO179 = ""; //預留欄位
                                    x2.MO180 = ""; //預留欄位
                                    x2.MO181 = 0; //APS固定工時
                                    x2.MO182 = 0; //APS變動工時
                                    x2.MO183 = 0; //批次加工量
                                    x2.MO184 = ""; //預留欄位
                                    x2.MO185 = 0; //基準數量
                                    x2.MO186 = ""; //資源群組
                                    x2.MO187 = ""; //資源群組名稱
                                    x2.MO188 = ""; //機台代號
                                    x2.MO189 = ""; //機台名稱
                                    x2.MO190 = ""; //指定資源
                                    x2.MO191 = ""; //關鍵料號
                                    x2.MO192 = 0; //營業稅率
                                    x2.MO193 = ""; //Bin管理
                                    x2.MO194 = ""; //交易設限碼
                                    x2.MO195 = 0; //保固佔售價比率
                                    x2.MO196 = 0; //保固期數(月數)
                                    x2.MO197 = ""; //
                                    x2.MO198 = ""; //預留欄位
                                    x2.MO199 = ""; //預留欄位
                                    x2.MO500 = ""; //產品系列ID
                                    x2.MO501 = 0; //GROSS_DIE
                                    x2.MO502 = ""; //說明1
                                    x2.MO503 = ""; //說明2
                                    x2.MO504 = ""; //光罩層次
                                    x2.MO505 = ""; //進項稅別碼
                                    x2.MO506 = ""; //銷項稅別碼
                                    x2.MO507 = 0; //
                                    x2.MO508 = 0; //
                                    x2.MO545 = 0; //標準產出良率
                                    x2.MO550 = ""; //產品屬性
                                    x2.MO551 = ""; //Wafer型號
                                    x2.MO552 = ""; //性質
                                    x2.MO553 = ""; //版本
                                    x2.MO554 = ""; //作業代號
                                    x2.UDF01 = "";
                                    x2.UDF02 = "";
                                    x2.UDF03 = "";
                                    x2.UDF04 = "";
                                    x2.UDF05 = "";
                                    x2.UDF06 = 0;
                                    x2.UDF07 = 0;
                                    x2.UDF08 = 0;
                                    x2.UDF09 = 0;
                                    x2.UDF10 = 0;
                                    #endregion
                                    #region //INSERT INTO INVMO
                                    sql = @"INSERT INTO INVMO (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , MO001, MO002, MO003, MO004, MO005, MO006, MO007, MO008, MO009, MO010
                                            , MO011, MO012, MO013, MO014, MO015, MO016, MO017, MO018, MO019, MO020
                                            , MO021, MO022, MO023, MO024, MO025, MO026, MO027, MO028, MO029, MO030
                                            , MO031, MO032, MO033, MO034, MO035, MO036, MO037, MO038, MO039, MO040
                                            , MO041, MO042, MO043, MO044, MO045, MO046, MO047, MO048, MO049, MO050
                                            , MO051, MO052, MO053, MO054, MO055, MO056, MO057, MO058, MO059, MO060
                                            , MO061, MO062, MO063, MO064, MO065, MO066, MO067, MO068, MO069, MO070
                                            , MO071, MO072, MO073, MO074, MO075, MO076, MO077, MO078, MO079, MO080
                                            , MO081, MO082, MO083, MO084, MO085, MO086, MO087, MO088, MO089, MO090
                                            , MO091, MO092, MO093, MO094, MO095, MO096, MO097, MO098, MO099, MO100
                                            , MO101, MO102, MO103, MO104, MO105, MO106, MO107, MO108, MO109, MO110
                                            , MO111, MO112, MO113, MO114, MO115, MO116, MO117, MO118, MO119, MO120
                                            , MO121, MO122, MO123, MO124, MO125, MO126, MO127, MO128, MO129, MO130
                                            , MO131, MO132, MO133, MO134, MO135, MO136, MO137, MO138, MO139, MO140
                                            , MO141, MO142, MO143, MO144, MO145, MO146, MO147, MO148, MO149, MO150
                                            , MO151, MO152, MO153, MO154, MO155, MO156, MO157, MO158, MO159, MO160
                                            , MO161, MO162, MO163, MO164, MO165, MO166, MO167, MO168, MO169, MO170
                                            , MO171, MO172, MO173, MO174, MO175, MO176, MO177, MO178, MO179, MO180
                                            , MO181, MO182, MO183, MO184, MO185, MO186, MO187, MO188, MO189, MO190
                                            , MO191, MO192, MO193, MO194, MO195, MO196, MO197, MO198, MO199, MO500
                                            , MO501, MO502, MO503, MO504, MO505, MO506, MO507, MO508
                                            , MO545, MO550, MO551, MO552, MO553, MO554
                                            , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @MO001, @MO002, @MO003, @MO004, @MO005, @MO006, @MO007, @MO008, @MO009, @MO010
                                            , @MO011, @MO012, @MO013, @MO014, @MO015, @MO016, @MO017, @MO018, @MO019, @MO020
                                            , @MO021, @MO022, @MO023, @MO024, @MO025, @MO026, @MO027, @MO028, @MO029, @MO030
                                            , @MO031, @MO032, @MO033, @MO034, @MO035, @MO036, @MO037, @MO038, @MO039, @MO040
                                            , @MO041, @MO042, @MO043, @MO044, @MO045, @MO046, @MO047, @MO048, @MO049, @MO050
                                            , @MO051, @MO052, @MO053, @MO054, @MO055, @MO056, @MO057, @MO058, @MO059, @MO060
                                            , @MO061, @MO062, @MO063, @MO064, @MO065, @MO066, @MO067, @MO068, @MO069, @MO070
                                            , @MO071, @MO072, @MO073, @MO074, @MO075, @MO076, @MO077, @MO078, @MO079, @MO080
                                            , @MO081, @MO082, @MO083, @MO084, @MO085, @MO086, @MO087, @MO088, @MO089, @MO090
                                            , @MO091, @MO092, @MO093, @MO094, @MO095, @MO096, @MO097, @MO098, @MO099, @MO100
                                            , @MO101, @MO102, @MO103, @MO104, @MO105, @MO106, @MO107, @MO108, @MO109, @MO110
                                            , @MO111, @MO112, @MO113, @MO114, @MO115, @MO116, @MO117, @MO118, @MO119, @MO120
                                            , @MO121, @MO122, @MO123, @MO124, @MO125, @MO126, @MO127, @MO128, @MO129, @MO130
                                            , @MO131, @MO132, @MO133, @MO134, @MO135, @MO136, @MO137, @MO138, @MO139, @MO140
                                            , @MO141, @MO142, @MO143, @MO144, @MO145, @MO146, @MO147, @MO148, @MO149, @MO150
                                            , @MO151, @MO152, @MO153, @MO154, @MO155, @MO156, @MO157, @MO158, @MO159, @MO160
                                            , @MO161, @MO162, @MO163, @MO164, @MO165, @MO166, @MO167, @MO168, @MO169, @MO170
                                            , @MO171, @MO172, @MO173, @MO174, @MO175, @MO176, @MO177, @MO178, @MO179, @MO180
                                            , @MO181, @MO182, @MO183, @MO184, @MO185, @MO186, @MO187, @MO188, @MO189, @MO190
                                            , @MO191, @MO192, @MO193, @MO194, @MO195, @MO196, @MO197, @MO198, @MO199, @MO500
                                            , @MO501, @MO502, @MO503, @MO504, @MO505, @MO506, @MO507, @MO508
                                            , @MO545, @MO550, @MO551, @MO552, @MO553, @MO554
                                            , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                    rowsAffected += erpConnection.Execute(sql, x2);
                                    #endregion
                                    break;
                                case "Y":
                                default:
                                    throw new SystemException("品號拋轉狀態錯誤!");
                            }

                            BOMMC x3 = billOfMaterials.FirstOrDefault(f => f.MC001 == x.MB001);
                            if (x3 != null)
                            {
                                BOMMC bommc = resultBOMMC.FirstOrDefault(f => f.MC001 == x3.MC001);
                                if (bommc == null)
                                {
                                    #region //主件新增 BOMMC from billOfMaterials
                                    x3.MC003 = "";
                                    x3.MC011 = 0;
                                    x3.MC012 = 0;
                                    x3.MC013 = "";
                                    x3.MC014 = "";
                                    x3.MC015 = "";
                                    x3.MC016 = "Y";
                                    x3.MC017 = dateNow;
                                    x3.MC019 = "N";
                                    x3.MC020 = 0;
                                    x3.MC021 = 0;
                                    x3.UDF01 = "";
                                    x3.UDF02 = "";
                                    x3.UDF03 = "";
                                    x3.UDF04 = "";
                                    x3.UDF05 = "";
                                    x3.UDF06 = 0;
                                    x3.UDF07 = 0;
                                    x3.UDF08 = 0;
                                    x3.UDF09 = 0;
                                    x3.UDF10 = 0;
                                    #endregion
                                    #region //INSERT INTO BOMMC
                                    sql = @"INSERT INTO BOMMC (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , MC001, MC002, MC003, MC004, MC005, MC006, MC007, MC008, MC009, MC010
                                            , MC011, MC012, MC013, MC014, MC015, MC016, MC017, MC018, MC019, MC020
                                            , MC021
                                            , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @MC001, @MC002, @MC003, @MC004, @MC005, @MC006, @MC007, @MC008, @MC009, @MC010
                                            , @MC011, @MC012, @MC013, @MC014, @MC015, @MC016, @MC017, @MC018, @MC019, @MC020
                                            , @MC021
                                            , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                    rowsAffected += erpConnection.Execute(sql, x3);
                                    #endregion
                                }
                                else { /*更新處理(由ERP代為編輯處理)*/ }

                                BOMMD x4 = bomDetails.FirstOrDefault(f => f.MD001 == x3.MC001);
                                if (x4 != null)
                                {
                                    BOMMD bommd = resultBOMMD.FirstOrDefault(f => f.MD001 == x4.MD001 && f.MD003 == x4.MD003);
                                    if (bommd == null)
                                    {
                                        #region //元件新增 BOMMD from bomDetail
                                        x4.MD005 = "";
                                        x4.MD009 = "****";
                                        x4.MD010 = "1";
                                        x4.MD013 = "N";
                                        x4.MD015 = "";
                                        x4.MD018 = 0;
                                        x4.MD019 = "";
                                        x4.MD020 = "";
                                        x4.MD021 = "";
                                        x4.MD022 = "";
                                        x4.MD023 = "";
                                        x4.MD024 = "";
                                        x4.MD025 = "";
                                        x4.MD026 = "";
                                        x4.MD027 = 0;
                                        x4.MD028 = 0;
                                        x4.MD029 = "Y";
                                        x4.MD030 = "";
                                        x4.MD031 = "";
                                        x4.MD032 = "Y";
                                        //x.UDF01 = "";
                                        x4.UDF02 = "";
                                        x4.UDF03 = "";
                                        x4.UDF04 = "";
                                        x4.UDF05 = "";
                                        x4.UDF06 = 0;
                                        x4.UDF07 = 0;
                                        x4.UDF08 = 0;
                                        x4.UDF09 = 0;
                                        x4.UDF10 = 0;
                                        #endregion
                                        #region //INSERT INTO BOMMD
                                        sql = @"INSERT INTO BOMMD (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , MD001, MD002, MD003, MD004, MD005, MD006, MD007, MD008, MD009, MD010
                                                , MD011, MD012, MD013, MD014, MD015, MD016, MD017, MD018, MD019, MD020
                                                , MD021, MD022, MD023, MD024, MD025, MD026, MD027, MD028, MD029, MD030
                                                , MD031, MD032
                                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                                VALUES ( @COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @MD001, @MD002, @MD003, @MD004, @MD005, @MD006, @MD007, @MD008, @MD009, @MD010
                                                , @MD011, @MD012, @MD013, @MD014, @MD015, @MD016, @MD017, @MD018, @MD019, @MD020
                                                , @MD021, @MD022, @MD023, @MD024, @MD025, @MD026, @MD027, @MD028, @MD029, @MD030
                                                , @MD031, @MD032
                                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                        rowsAffected += erpConnection.Execute(sql, x4);
                                        #endregion
                                    }
                                    else { /*更新處理(由ERP代為編輯處理)*/ }
                                }
                            }
                        });
                    }

                    //後處理
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //回寫品號資料、拋轉狀態和時間
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.TransferStatus = @TransferStatus,
                                a.TransferDate = @TransferDate,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM PDM.MtlItem a
                                WHERE MtlItemId IN @MtlItemIds";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatus = "Y",
                                TransferDate = dateNow,
                                LastModifiedDate,
                                LastModifiedBy,
                                MtlItemIds = mtlItems.Select(s => s.MtlItemId).Distinct().ToArray()
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //判斷品號是否拋轉完成
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MtlItemNo, TransferStatus
                                FROM PDM.MtlItem
                                WHERE MtlItemId IN @MtlItemIds";
                        dynamicParameters.Add("MtlItemIds", mtlItems.Select(s => s.MtlItemId).Distinct().ToArray());
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (!result.Any()) throw new SystemException("【品號】資料錯誤!");
                        if (result.Any(a => a.TransferStatus == "N")) throw new SystemException("【品號】拋轉未完成!");
                        #endregion

                        #region//回寫主件確認狀態、拋轉狀態及時間
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.ConfirmStatus = @ConfirmStatus,
                                a.ConfirmUserId = @ConfirmUserId,
                                a.ConfirmDate = @ConfirmDate,
                                a.TransferStatus = @TransferStatus,
                                a.TransferDate = @TransferDate,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy 
                                FROM PDM.BillOfMaterials a
                                WHERE MtlItemId IN @MtlItemIds";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUserId = CurrentUser,
                                ConfirmDate = dateNow,
                                TransferStatus = "Y",
                                TransferDate = dateNow,
                                LastModifiedDate,
                                LastModifiedBy,
                                MtlItemIds = mtlItems.Select(s => s.MtlItemId).Distinct().ToArray()
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region//回寫元件資料、拋轉狀態和時間
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET
                                a.ConfirmStatus = @ConfirmStatus,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                FROM PDM.BomDetail a
                                INNER JOIN PDM.BillOfMaterials b ON b.BomId = a.BomId 
                                WHERE b.MtlItemId IN @MtlItemIds";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                MtlItemIds = mtlItems.Select(s => s.MtlItemId).Distinct().ToArray()
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
        #region //UpdateBillOfMaterials -- Bom主件資料更新 -- Zoey 2022.12.06
        public string UpdateBillOfMaterials(int BomId, int MtlItemId, double StandardLot, string WipPrefix, string Remark)
        {
            try
            {
                if (StandardLot <= 0) throw new SystemException("【標準批量】不能為空!");
                if (WipPrefix.Length <= 0) throw new SystemException("【製令單別】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        string TransferStatus = "";

                        #region //判斷Bom主件是否存在/撈取拋轉狀態
                        sql = @"SELECT TransferStatus, MtlItemId
                                FROM PDM.BillOfMaterials
                                WHERE BomId = @BomId";
                        dynamicParameters.Add("BomId", BomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in result)
                        {
                            TransferStatus = item.TransferStatus;
                            MtlItemId = item.MtlItemId;
                        }
                        #endregion

                        if (TransferStatus != "Y")
                        {
                            if (result.Count() <= 0)
                            {
                                #region //判斷品號是否正確
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.MtlItem
                                        WHERE MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() <= 0) throw new SystemException("【品號】資料錯誤!");
                                #endregion

                                #region //撈取單位ID
                                int uomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT UomId FROM PDM.UnitOfMeasure
                                        WHERE UomNo = 'PCS'
                                        AND CompanyId = @CompanyId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var result3 = sqlConnection.Query(sql, dynamicParameters);

                                if (result3.Count() <= 0) throw new SystemException("單位資料錯誤!");

                                foreach (var item in result3)
                                {
                                    uomId = item.UomId;
                                }
                                #endregion

                                #region //Bom主件新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.BillOfMaterials (MtlItemId, UomId, StandardLot, WipPrefix, ModiPrefix, ModiNo, ModiSequence, Version
                                        , Remark, ConfirmStatus, ConfirmUserId, ConfirmDate, TransferStatus, TransferDate, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.BomId
                                        VALUES (@MtlItemId, @UomId, @StandardLot, @WipPrefix, @ModiPrefix, @ModiNo, @ModiSequence, @Version
                                        , @Remark, @ConfirmStatus, @ConfirmUserId, @ConfirmDate, @TransferStatus, @TransferDate, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MtlItemId,
                                        UomId = uomId,
                                        StandardLot,
                                        WipPrefix,
                                        ModiPrefix = "",
                                        ModiNo = "",
                                        ModiSequence = "",
                                        Version = "0000",
                                        Remark = "",
                                        ConfirmStatus = 'N',
                                        ConfirmUserId = (int?)null,
                                        ConfirmDate = (int?)null,
                                        TransferStatus = 'N',
                                        TransferDate = (int?)null,
                                        Status = 'Y',
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected = insertResult.Count();
                                #endregion
                            }
                            else
                            {
                                #region //Bom主件修改
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.BillOfMaterials SET
                                        StandardLot = @StandardLot,
                                        WipPrefix = @WipPrefix,
                                        Remark = @Remark,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE BomId = @BomId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        StandardLot,
                                        WipPrefix,
                                        Remark,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        BomId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                        }
                        else
                        {
                            throw new SystemException("BOM已拋轉，無法增修!");
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

        #region //UpdateBomDetail -- Bom元件資料更新 -- Zoey 2022.12.06
        public string UpdateBomDetail(int BomDetailId, int BomId, int MtlItemId, string BomSequence, double CompositionQuantity, double Base, double LossRate
            , string StandardCostingType, string MaterialProperties, string EffectiveDate, string ExpirationDate, string Remark)
        {
            try
            {
                if (MtlItemId <= 0) throw new SystemException("【元件品號】不能為空!");
                if (BomSequence.Length <= 0) throw new SystemException("【序號】不能為空!");
                if (BomSequence.Length > 4) throw new SystemException("【序號】長度錯誤!");
                if (CompositionQuantity < 0) throw new SystemException("【單位用量】不能為空!");
                if (Base < 0) throw new SystemException("【底數】不能為空!");
                if (LossRate < 0) throw new SystemException("【損壞率】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (StandardCostingType != "Y" && StandardCostingType != "N") throw new SystemException("【標準成本計算】不能為空!");
                if (MaterialProperties != "1" && MaterialProperties != "2" && MaterialProperties != "3" && MaterialProperties != "4" && MaterialProperties != "5") throw new SystemException("【材料性質】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        string TransferStatus = "";
                        int rowsAffected = 0;

                        #region //判斷Bom元件是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.BomDetail
                                WHERE BomDetailId = @BomDetailId";
                        dynamicParameters.Add("BomDetailId", BomDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("Bom元件資料錯誤!");
                        #endregion

                        #region //判斷元件品號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.BomDetail
                                WHERE MtlItemId = @MtlItemId
                                AND BomId = @BomId
                                AND BomDetailId != @BomDetailId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);
                        dynamicParameters.Add("BomId", BomId);
                        dynamicParameters.Add("BomDetailId", BomDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【元件品號】重複，請重新輸入!");
                        #endregion

                        #region //撈取拋轉狀態
                        sql = @"SELECT TransferStatus
                                FROM PDM.BillOfMaterials
                                WHERE BomId = @BomId";
                        dynamicParameters.Add("BomId", BomId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in result3)
                        {
                            TransferStatus = item.TransferStatus;
                        }
                        #endregion

                        if (TransferStatus != "Y")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.BomDetail SET
                                    MtlItemId = @MtlItemId,
                                    BomSequence = @BomSequence,
                                    CompositionQuantity = @CompositionQuantity,
                                    Base = @Base,
                                    LossRate = @LossRate,
                                    StandardCostingType=@StandardCostingType,
                                    MaterialProperties=@MaterialProperties,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BomDetailId = @BomDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MtlItemId,
                                    BomSequence,
                                    CompositionQuantity,
                                    Base,
                                    LossRate,
                                    StandardCostingType,
                                    MaterialProperties,
                                    Remark,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    BomDetailId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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
    }
}
