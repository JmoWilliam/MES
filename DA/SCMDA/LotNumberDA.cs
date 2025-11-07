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
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class LotNumberDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string ErpSysDbConnectionStrings = "";
        public string BpmDbConnectionStrings = "";

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

        public LotNumberDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            ErpSysDbConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
            BpmDbConnectionStrings = ConfigurationManager.AppSettings["BpmDb"];

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

        #region //Get
        #region //GetLotNumber -- 取得批號資料 -- Ann 2024-03-20
        public string GetLotNumber(int LotNumberId, string MtlItemNo, string MtlItemName, string LotNumberNo, string CloseStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.LotNumberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MtlItemId, a.LotNumberNo, FORMAT(a.FirstRecriptDate, 'yyyy-MM-dd') FirstRecriptDate
                        , a.FromErpPrefix, a.FromErpNo, a.CloseStatus, a.Remark, a.TransferStatus
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                        , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec";
                    sqlQuery.mainTables =
                        @"FROM SCM.LotNumber a 
                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND b.MtlItemNo = @MtlItemNo", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LotNumberId", @" AND a.LotNumberId = @LotNumberId", LotNumberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND b.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND b.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LotNumberNo", @" AND a.LotNumberNo LIKE '%' + @LotNumberNo + '%'", LotNumberNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CloseStatus", @" AND a.CloseStatus = @CloseStatus", CloseStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LotNumberId DESC";
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

        #region //GetLnDetail -- 取得批號詳細資料 -- Ann 2024-03-22
        public string GetLnDetail(int LnDetailId, int LotNumberId, string MtlItemNo, string MtlItemName, string FromErpFullNo, int InventoryId, int TransactionType, string DocType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.LnDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.LotNumberId, FORMAT(a.TransactionDate, 'yyyy-MM-dd') TransactionDate, a.FromErpPrefix, a.FromErpNo, a.FromSeq
                        , a.InventoryId, a.TransactionType, a.DocType, a.Quantity, a.Remark
                        , c.InventoryNo, c.InventoryName
                        , d.MtlItemNo, d.MtlItemName, d.MtlItemSpec
                        , e.TypeName DocTypeNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.LnDetail a 
                        INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                        INNER JOIN PDM.MtlItem d ON b.MtlItemId = d.MtlItemId
                        INNER JOIN BAS.[Type] e ON a.DocType = e.TypeNo AND e.TypeSchema = 'LnDetail.DocType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LnDetailId", @" AND a.LnDetailId = @LnDetailId", LnDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LotNumberId", @" AND a.LotNumberId = @LotNumberId", LotNumberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND d.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND d.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FromErpFullNo", @" AND (a.FromErpPrefix + '-' + b.FromErpNo +  '(' + CONVERT(VARCHAR(10), a.FromSeq) + ')') LIKE '%' + @FromErpFullNo + '%'", FromErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryId", @" AND a.InventoryId = @InventoryId", InventoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransactionType", @" AND a.TransactionType = @TransactionType", TransactionType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DocType", @" AND a.DocType = @DocType", DocType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LnDetailId DESC";
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
        #endregion

        #region //Add
        #region //AddLotNumber -- 新增批號資料 -- Ann 2024-03-21
        public string AddLotNumber(int MtlItemId, string LotNumberNo , string Remark)
        {
            try
            {
                if (LotNumberNo.Length <= 0) throw new SystemException("【批號】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo, a.LotManagement
                                    FROM PDM.MtlItem a 
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                            string MtlItemNo = "";
                            foreach (var item in MtlItemResult)
                            {
                                if (item.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                        , LTRIM(RTRIM(MB030)) MB030
                                        , LTRIM(RTRIM(MB031)) MB031
                                        , LTRIM(RTRIM(MB022)) MB022
                                        , LTRIM(RTRIM(MB043)) MB043
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item2 in INVMBResult)
                            {
                                #region //判斷ERP品號生效日與失效日
                                if (item2.MB030 != "" && item2.MB030 != null)
                                {
                                    #region //判斷單據日期需大於或等於生效日
                                    string EffectiveDate = item2.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DateTime.Now, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
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
                                    int effresult = DateTime.Compare(DateTime.Now, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                                #endregion

                                #region //確認此品號是否需要批號管理
                                if (item2.MB022 != "T")
                                {
                                    throw new SystemException("品號【" + MtlItemNo + "】批號控管參數錯誤!!");
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認此品號是否已存在此批號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            dynamicParameters.Add("LotNumberNo", LotNumberNo);

                            var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                            if (LotNumberResult.Count() > 0) throw new SystemException("品號【" + MtlItemNo + "】已存在此批號【" + LotNumberNo + "】!");
                            #endregion

                            #region //INSERT SCM.GoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.LotNumber (MtlItemId, LotNumberNo, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LotNumberId
                                    VALUES (@MtlItemId, @LotNumberNo, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MtlItemId,
                                    LotNumberNo,
                                    Remark,
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

        #region //Update
        #region //UpdateLotNumberSynchronize -- 批號資料同步 -- Ann 2024-03-18
        public string UpdateLotNumberSynchronize(string CompanyNo, string UpdateDate, string NormalSync, string TranSync)
        {
            try
            {
                List<LotNumber> lotNumbers = new List<LotNumber>();
                List<LnDetail> lnDetails = new List<LnDetail>();

                if (NormalSync == null) NormalSync = "N";
                if (TranSync == null) TranSync = "Y";

                int rowsAffected = 0;
                int mainAffected = 0;
                int detailAffected = 0;
                int mainDelAffected = 0;
                int detailDelAffected = 0;

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

                    #region //正常同步
                    if (TranSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //撈取ERP批號單頭資料
                            sql = @"SELECT LTRIM(RTRIM(a.ME001)) MtlItemNo, LTRIM(RTRIM(a.ME002)) LotNumberNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(a.ME003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(a.ME003)) as date), 'yyyy-MM-dd') ELSE NULL END FirstRecriptDate
                                    , LTRIM(RTRIM(a.ME005)) FromErpPrefix, LTRIM(RTRIM(a.ME006)) FromErpNo, LTRIM(RTRIM(a.ME007)) CloseStatus
                                    , LTRIM(RTRIM(a.ME008)) Remark
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVME a 
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                            sql += @" ORDER BY TransferDate, TransferTime";

                            lotNumbers = sqlConnection.Query<LotNumber>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP批號單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF001)) MtlItemNo, LTRIM(RTRIM(a.MF002)) LotNumberNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(a.MF003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(a.MF003)) as date), 'yyyy-MM-dd') ELSE NULL END TransactionDate
                                    , LTRIM(RTRIM(a.MF004)) FromErpPrefix, LTRIM(RTRIM(a.MF005)) FromErpNo, LTRIM(RTRIM(a.MF006)) FromSeq, LTRIM(RTRIM(a.MF007)) InventoryNo
                                    , a.MF008 TransactionType, LTRIM(RTRIM(a.MF009)) DocType, a.MF010 Quantity, LTRIM(RTRIM(a.MF013)) Remark
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVMF a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                            sql += @" ORDER BY TransferDate, TransferTime";

                            lnDetails = sqlConnection.Query<LnDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            lotNumbers = lotNumbers.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            lnDetails = lnDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.LotNumber是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LotNumberId, a.MtlItemId, a.LotNumberNo
                                    FROM SCM.LotNumber a 
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<LotNumber> resultLotNumber = sqlConnection.Query<LotNumber>(sql, dynamicParameters).ToList();

                            lotNumbers = lotNumbers.GroupJoin(resultLotNumber, x => new { x.MtlItemId, x.LotNumberNo }, y => new { y.MtlItemId, y.LotNumberNo }, (x, y) => { x.LotNumberId = y.FirstOrDefault()?.LotNumberId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //批號單頭(新增/修改)
                            List<LotNumber> addLotNumbers = lotNumbers.Where(x => x.LotNumberId == null).ToList();
                            List<LotNumber> updateLotNumbers = lotNumbers.Where(x => x.LotNumberId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addLotNumbers.Count > 0)
                            {
                                addLotNumbers
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.LotNumber (MtlItemId, LotNumberNo, FirstRecriptDate, FromErpPrefix, FromErpNo, CloseStatus, Remark, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@MtlItemId, @LotNumberNo, @FirstRecriptDate, @FromErpPrefix, @FromErpNo, @CloseStatus, @Remark, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                mainAffected += sqlConnection.Execute(sql, addLotNumbers);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateLotNumbers.Count > 0)
                            {
                                updateLotNumbers
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.LotNumber SET
                                        MtlItemId = @MtlItemId,
                                        LotNumberNo = @LotNumberNo,
                                        FirstRecriptDate = @FirstRecriptDate,
                                        FromErpPrefix = @FromErpPrefix,
                                        FromErpNo = @FromErpNo,
                                        CloseStatus = @CloseStatus,
                                        Remark = @Remark,
                                        TransferStatus = @TransferStatus,
                                        TransferDate = @TransferDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE LotNumberId = @LotNumberId";
                                mainAffected += sqlConnection.Execute(sql, updateLotNumbers);
                            }
                            #endregion
                            #endregion

                            #region //批號單身(新增/修改)
                            #region //撈取批號單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LotNumberId, a.MtlItemId, a.LotNumberNo
                                    , b.MtlItemNo
                                    FROM SCM.LotNumber a 
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            resultLotNumber = sqlConnection.Query<LotNumber>(sql, dynamicParameters).ToList();

                            lnDetails = lnDetails.Join(resultLotNumber, x => new { x.MtlItemNo, x.LotNumberNo }, y => new { y.MtlItemNo, y.LotNumberNo }, (x, y) => { x.LotNumberId = y.LotNumberId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.LnDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LnDetailId, a.LotNumberId, a.FromErpPrefix, a.FromErpNo, a.FromSeq, a.InventoryId
                                    FROM SCM.LnDetail a
                                    INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    WHERE c.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<LnDetail> resultLnDetail = sqlConnection.Query<LnDetail>(sql, dynamicParameters).ToList();

                            lnDetails = lnDetails.GroupJoin(resultLnDetail, x => new { x.LotNumberId, x.FromErpPrefix, x.FromErpNo, x.FromSeq, x.InventoryId }, y => new { y.LotNumberId, y.FromErpPrefix, y.FromErpNo, y.FromSeq, y.InventoryId }, (x, y) => { x.LnDetailId = y.FirstOrDefault()?.LnDetailId; return x; }).ToList();
                            #endregion

                            List<LnDetail> addLnDetail = lnDetails.Where(x => x.LnDetailId == null).ToList();
                            List<LnDetail> updateLnDetail = lnDetails.Where(x => x.LnDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addLnDetail.Count > 0)
                            {
                                addLnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix, FromErpNo, FromSeq, InventoryId, TransactionType
                                        , DocType, Quantity, Remark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix, @FromErpNo, @FromSeq, @InventoryId, @TransactionType
                                        , @DocType, @Quantity, @Remark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                detailAffected += sqlConnection.Execute(sql, addLnDetail);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateLnDetail.Count > 0)
                            {
                                updateLnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.LnDetail SET
                                    LotNumberId = @LotNumberId,
                                    TransactionDate = @TransactionDate,
                                    FromErpPrefix = @FromErpPrefix,
                                    FromErpNo = @FromErpNo,
                                    FromSeq = @FromSeq,
                                    InventoryId = @InventoryId,
                                    TransactionType = @TransactionType,
                                    DocType = @DocType,
                                    Quantity = @Quantity,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE LnDetailId = @LnDetailId";
                                detailAffected += sqlConnection.Execute(sql, updateLnDetail);
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion

                    #region //異動同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP批號單頭資料
                            sql = @"SELECT LTRIM(RTRIM(ME001)) MtlItemNo, LTRIM(RTRIM(ME002)) LotNumberNo
                                    FROM INVME
                                    WHERE 1=1
                                    ORDER BY ME001, ME002";
                            var resultErpLn = erpConnection.Query<LotNumber>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM批號單頭資料
                                sql = @"SELECT a.MtlItemId, a.LotNumberNo
                                        , b.MtlItemNo
                                        FROM SCM.LotNumber a 
                                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.MtlItemNo, a.LotNumberNo";
                                dynamicParameters.Add("CompanyId", CompanyId);
                                var resultBmLn = bmConnection.Query<LotNumber>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的批號單頭
                                var dictionaryErpLn = resultErpLn.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo, x => x);
                                var dictionaryBmLn = resultBmLn.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo, x => x);
                                var changeLn = dictionaryBmLn.Where(x => !dictionaryErpLn.ContainsKey(x.Key)).ToList();
                                var changeLnList = changeLn.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動訂單單頭
                                if (changeLnList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.LnDetail
                                            WHERE LotNumberId IN (
                                                SELECT a.LotNumberId
                                                FROM SCM.LotNumber a 
                                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                                WHERE b.MtlItemNo + '_' + a.LotNumberNo IN @LotNumberNo
                                            )";
                                    dynamicParameters.Add("LotNumberNo", changeLnList.Select(x => x.MtlItemNo + "_" + x.LotNumberNo).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a 
                                            FROM SCM.LotNumber a 
                                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                            WHERE b.MtlItemNo + '_' + a.LotNumberNo IN @LotNumberNo";
                                    dynamicParameters.Add("LotNumberNo", changeLnList.Select(x => x.MtlItemNo + "_" + x.LotNumberNo).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP批號單身資料
                            sql = @"SELECT LTRIM(RTRIM(MF001)) MtlItemNo, LTRIM(RTRIM(MF002)) LotNumberNo, LTRIM(RTRIM(MF004)) FromErpPrefix
                                    , LTRIM(RTRIM(MF005)) FromErpNo, LTRIM(RTRIM(MF006)) FromSeq, LTRIM(RTRIM(MF007)) InventoryNo
                                    FROM INVMF
                                    WHERE 1=1
                                    ORDER BY MF001, MF002";
                            var resultErpLnDetail = erpConnection.Query<LnDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM批號單身資料
                                sql = @"SELECT c.MtlItemNo, b.LotNumberNo, a.FromErpPrefix, a.FromErpNo, a.FromSeq, d.InventoryNo
                                        FROM SCM.LnDetail a 
                                        INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                        INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                        INNER JOIN SCM.Inventory d ON a.InventoryId = d.InventoryId
                                        WHERE c.CompanyId = @CompanyId
                                        ORDER BY c.MtlItemNo, b.LotNumberNo";
                                dynamicParameters.Add("CompanyId", CompanyId);
                                var resultBmLnDetail = bmConnection.Query<LnDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的批號單身
                                var dictionaryErpLnDetail = resultErpLnDetail.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo, x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo);
                                var dictionaryBmLnDetail = resultBmLnDetail.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo, x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo);
                                var changeLnDetail = dictionaryBmLnDetail.Where(x => !dictionaryErpLnDetail.ContainsKey(x.Key)).ToList();
                                var changeLnDetailList = changeLnDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除訂單單身
                                if (changeLnDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.LnDetail a
                                            INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                            INNER JOIN SCM.Inventory d ON a.InventoryId = d.InventoryId
                                            WHERE c.MtlItemNo + '_' + b.LotNumberNo + '_' + a.FromErpPrefix + '_' + a.FromErpNo + '_' + a.FromSeq + '_' + d.InventoryNo IN @LnDeatilInfo";
                                    dynamicParameters.Add("LnDeatilInfo", changeLnDetailList.Select(x => x).ToArray());
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
                        data = "已更新資料【" + mainAffected + "】筆單頭, 【" + detailAffected + "】筆單身, 刪除【" + mainDelAffected + "】筆單頭, 【" + detailDelAffected + "】筆單身"
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

        #region //UpdateLotNumber -- 更新批號資料 -- Ann 2024-03-21
        public string UpdateLotNumber(int LotNumberId, int MtlItemId, string LotNumberNo, string Remark)
        {
            try
            {
                if (LotNumberNo.Length <= 0) throw new SystemException("【批號】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認批號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.LotNumber a 
                                    WHERE a.LotNumberId = @LotNumberId";
                            dynamicParameters.Add("LotNumberId", LotNumberId);

                            var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                            if (LotNumberResult.Count() <= 0) throw new SystemException("批號資料錯誤!!");
                            #endregion

                            #region //確認此批號無任何單身交易紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1 
                                    FROM SCM.LnDetail a 
                                    WHERE a.LotNumberId = @LotNumberId";
                            dynamicParameters.Add("LotNumberId", LotNumberId);

                            var LnDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (LnDetailResult.Count() > 0) throw new SystemException("此批號已經有交易紀錄，無法更改!!");
                            #endregion

                            #region //確認品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo, a.LotManagement
                                    FROM PDM.MtlItem a 
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                            string MtlItemNo = "";
                            foreach (var item in MtlItemResult)
                            {
                                if (item.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                        , LTRIM(RTRIM(MB030)) MB030
                                        , LTRIM(RTRIM(MB031)) MB031
                                        , LTRIM(RTRIM(MB022)) MB022
                                        , LTRIM(RTRIM(MB043)) MB043
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item2 in INVMBResult)
                            {
                                #region //判斷ERP品號生效日與失效日
                                if (item2.MB030 != "" && item2.MB030 != null)
                                {
                                    #region //判斷單據日期需大於或等於生效日
                                    string EffectiveDate = item2.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DateTime.Now, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
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
                                    int effresult = DateTime.Compare(DateTime.Now, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                                #endregion

                                #region //確認此品號是否需要批號管理
                                if (item2.MB022 != "T")
                                {
                                    throw new SystemException("品號【" + MtlItemNo + "】批號控管參數錯誤!!");
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認此品號是否已存在此批號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo
                                    AND a.LotNumberId != @LotNumberId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            dynamicParameters.Add("LotNumberNo", LotNumberNo);
                            dynamicParameters.Add("LotNumberId", LotNumberId);

                            var LotNumberResult2 = sqlConnection.Query(sql, dynamicParameters);

                            if (LotNumberResult2.Count() > 0) throw new SystemException("品號【" + MtlItemNo + "】已存在此批號【" + LotNumberNo + "】!");
                            #endregion

                            #region //UPDATE SCM.GoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.LotNumber SET
                                    MtlItemId = @MtlItemId,
                                    LotNumberNo = @LotNumberNo,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE LotNumberId = @LotNumberId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                LotNumberNo,
                                Remark,
                                LastModifiedDate,
                                LastModifiedBy,
                                LotNumberId
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

        #region //UpdateLotNumberManualSynchronize -- 批號手動資料同步 -- Ann 2024-03-26
        public string UpdateLotNumberManualSynchronize(string MtlItemNo, string LotNumber, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync)
        {
            try
            {
                List<LotNumber> lotNumbers = new List<LotNumber>();
                List<LnDetail> lnDetails = new List<LnDetail>();

                int rowsAffected = 0;
                int mainAffected = 0;
                int detailAffected = 0;
                int mainDelAffected = 0;
                int detailDelAffected = 0;

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
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    #region //正常同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //撈取ERP批號單頭資料
                            sql = @"SELECT LTRIM(RTRIM(a.ME001)) MtlItemNo, LTRIM(RTRIM(a.ME002)) LotNumberNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(a.ME003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(a.ME003)) as date), 'yyyy-MM-dd') ELSE NULL END FirstRecriptDate
                                    , LTRIM(RTRIM(a.ME005)) FromErpPrefix, LTRIM(RTRIM(a.ME006)) FromErpNo, LTRIM(RTRIM(a.ME007)) CloseStatus
                                    , LTRIM(RTRIM(a.ME008)) Remark
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVME a 
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND LTRIM(RTRIM(a.ME001)) LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "LotNumber", @" AND LTRIM(RTRIM(a.ME002)) LIKE '%' + @LotNumber + '%'", LotNumber);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            lotNumbers = sqlConnection.Query<LotNumber>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP批號單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF001)) MtlItemNo, LTRIM(RTRIM(a.MF002)) LotNumberNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(a.MF003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(a.MF003)) as date), 'yyyy-MM-dd') ELSE NULL END TransactionDate
                                    , LTRIM(RTRIM(a.MF004)) FromErpPrefix, LTRIM(RTRIM(a.MF005)) FromErpNo, LTRIM(RTRIM(a.MF006)) FromSeq, LTRIM(RTRIM(a.MF007)) InventoryNo
                                    , a.MF008 TransactionType, LTRIM(RTRIM(a.MF009)) DocType, a.MF010 Quantity, LTRIM(RTRIM(a.MF013)) Remark
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVMF a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND LTRIM(RTRIM(a.MF001)) LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "LotNumber", @" AND LTRIM(RTRIM(a.MF002)) LIKE '%' + @LotNumber + '%'", LotNumber);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            lnDetails = sqlConnection.Query<LnDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            lotNumbers = lotNumbers.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            lnDetails = lnDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.LotNumber是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LotNumberId, a.MtlItemId, a.LotNumberNo
                                    FROM SCM.LotNumber a 
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<LotNumber> resultLotNumber = sqlConnection.Query<LotNumber>(sql, dynamicParameters).ToList();

                            lotNumbers = lotNumbers.GroupJoin(resultLotNumber, x => new { x.MtlItemId, x.LotNumberNo }, y => new { y.MtlItemId, y.LotNumberNo }, (x, y) => { x.LotNumberId = y.FirstOrDefault()?.LotNumberId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //批號單頭(新增/修改)
                            List<LotNumber> addLotNumbers = lotNumbers.Where(x => x.LotNumberId == null).ToList();
                            List<LotNumber> updateLotNumbers = lotNumbers.Where(x => x.LotNumberId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addLotNumbers.Count > 0)
                            {
                                addLotNumbers
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.LotNumber (MtlItemId, LotNumberNo, FirstRecriptDate, FromErpPrefix, FromErpNo, CloseStatus, Remark, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@MtlItemId, @LotNumberNo, @FirstRecriptDate, @FromErpPrefix, @FromErpNo, @CloseStatus, @Remark, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                mainAffected += sqlConnection.Execute(sql, addLotNumbers);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateLotNumbers.Count > 0)
                            {
                                updateLotNumbers
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.LotNumber SET
                                        MtlItemId = @MtlItemId,
                                        LotNumberNo = @LotNumberNo,
                                        FirstRecriptDate = @FirstRecriptDate,
                                        FromErpPrefix = @FromErpPrefix,
                                        FromErpNo = @FromErpNo,
                                        CloseStatus = @CloseStatus,
                                        Remark = @Remark,
                                        TransferStatus = @TransferStatus,
                                        TransferDate = @TransferDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE LotNumberId = @LotNumberId";
                                mainAffected += sqlConnection.Execute(sql, updateLotNumbers);
                            }
                            #endregion
                            #endregion

                            #region //批號單身(新增/修改)
                            #region //撈取批號單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LotNumberId, a.MtlItemId, a.LotNumberNo
                                    , b.MtlItemNo
                                    FROM SCM.LotNumber a 
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            resultLotNumber = sqlConnection.Query<LotNumber>(sql, dynamicParameters).ToList();

                            lnDetails = lnDetails.Join(resultLotNumber, x => new { x.MtlItemNo, x.LotNumberNo }, y => new { y.MtlItemNo, y.LotNumberNo }, (x, y) => { x.LotNumberId = y.LotNumberId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.LnDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LnDetailId, a.LotNumberId, a.FromErpPrefix, a.FromErpNo, a.FromSeq, a.InventoryId
                                    FROM SCM.LnDetail a
                                    INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    WHERE c.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CompanyId);

                            List<LnDetail> resultLnDetail = sqlConnection.Query<LnDetail>(sql, dynamicParameters).ToList();

                            lnDetails = lnDetails.GroupJoin(resultLnDetail, x => new { x.LotNumberId, x.FromErpPrefix, x.FromErpNo, x.FromSeq, x.InventoryId }, y => new { y.LotNumberId, y.FromErpPrefix, y.FromErpNo, y.FromSeq, y.InventoryId }, (x, y) => { x.LnDetailId = y.FirstOrDefault()?.LnDetailId; return x; }).ToList();
                            #endregion

                            List<LnDetail> addLnDetail = lnDetails.Where(x => x.LnDetailId == null).ToList();
                            List<LnDetail> updateLnDetail = lnDetails.Where(x => x.LnDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addLnDetail.Count > 0)
                            {
                                addLnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix, FromErpNo, FromSeq, InventoryId, TransactionType
                                        , DocType, Quantity, Remark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix, @FromErpNo, @FromSeq, @InventoryId, @TransactionType
                                        , @DocType, @Quantity, @Remark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                detailAffected += sqlConnection.Execute(sql, addLnDetail);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateLnDetail.Count > 0)
                            {
                                updateLnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.LnDetail SET
                                    LotNumberId = @LotNumberId,
                                    TransactionDate = @TransactionDate,
                                    FromErpPrefix = @FromErpPrefix,
                                    FromErpNo = @FromErpNo,
                                    FromSeq = @FromSeq,
                                    InventoryId = @InventoryId,
                                    TransactionType = @TransactionType,
                                    DocType = @DocType,
                                    Quantity = @Quantity,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE LnDetailId = @LnDetailId";
                                detailAffected += sqlConnection.Execute(sql, updateLnDetail);
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
                            #region //ERP批號單頭資料
                            sql = @"SELECT LTRIM(RTRIM(ME001)) MtlItemNo, LTRIM(RTRIM(ME002)) LotNumberNo
                                    FROM INVME
                                    WHERE 1=1
                                    ORDER BY ME001, ME002";
                            var resultErpLn = erpConnection.Query<LotNumber>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM批號單頭資料
                                sql = @"SELECT a.MtlItemId, a.LotNumberNo
                                        , b.MtlItemNo
                                        FROM SCM.LotNumber a 
                                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.MtlItemNo, a.LotNumberNo";
                                dynamicParameters.Add("CompanyId", CompanyId);
                                var resultBmLn = bmConnection.Query<LotNumber>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的批號單頭
                                var dictionaryErpLn = resultErpLn.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo, x => x);
                                var dictionaryBmLn = resultBmLn.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo, x => x);
                                var changeLn = dictionaryBmLn.Where(x => !dictionaryErpLn.ContainsKey(x.Key)).ToList();
                                var changeLnList = changeLn.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動訂單單頭
                                if (changeLnList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.LnDetail
                                            WHERE LotNumberId IN (
                                                SELECT a.LotNumberId
                                                FROM SCM.LotNumber a 
                                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                                WHERE b.MtlItemNo + '_' + a.LotNumberNo IN @LotNumberNo
                                            )";
                                    dynamicParameters.Add("LotNumberNo", changeLnList.Select(x => x.MtlItemNo + "_" + x.LotNumberNo).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a 
                                            FROM SCM.LotNumber a 
                                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                            WHERE b.MtlItemNo + '_' + a.LotNumberNo = @LotNumberNo";
                                    dynamicParameters.Add("LotNumberNo", changeLnList.Select(x => x.MtlItemNo + "_" + x.LotNumberNo).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP批號單身資料
                            sql = @"SELECT LTRIM(RTRIM(MF001)) MtlItemNo, LTRIM(RTRIM(MF002)) LotNumberNo, LTRIM(RTRIM(MF004)) FromErpPrefix
                                    , LTRIM(RTRIM(MF005)) FromErpNo, LTRIM(RTRIM(MF006)) FromSeq, LTRIM(RTRIM(MF007)) InventoryNo
                                    FROM INVMF
                                    WHERE 1=1
                                    ORDER BY MF001, MF002";
                            var resultErpLnDetail = erpConnection.Query<LnDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM批號單身資料
                                sql = @"SELECT c.MtlItemNo, b.LotNumberNo, a.FromErpPrefix, a.FromErpNo, a.FromSeq, d.InventoryNo
                                        FROM SCM.LnDetail a 
                                        INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                        INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                        INNER JOIN SCM.Inventory d ON a.InventoryId = d.InventoryId
                                        WHERE c.CompanyId = @CompanyId
                                        ORDER BY c.MtlItemNo, b.LotNumberNo";
                                dynamicParameters.Add("CompanyId", CompanyId);
                                var resultBmLnDetail = bmConnection.Query<LnDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的批號單身
                                var dictionaryErpLnDetail = resultErpLnDetail.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo, x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo);
                                var dictionaryBmLnDetail = resultBmLnDetail.ToDictionary(x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo, x => x.MtlItemNo + "_" + x.LotNumberNo + "_" + x.FromErpPrefix + "_" + x.FromErpNo + "_" + x.FromSeq + "_" + x.InventoryNo);
                                var changeLnDetail = dictionaryBmLnDetail.Where(x => !dictionaryErpLnDetail.ContainsKey(x.Key)).ToList();
                                var changeLnDetailList = changeLnDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除訂單單身
                                if (changeLnDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.LnDetail a
                                            INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                            INNER JOIN SCM.Inventory d ON a.InventoryId = d.InventoryId
                                            WHERE c.MtlItemNo + '_' + b.LotNumberNo + '_' + a.FromErpPrefix + '_' + a.FromErpNo + '_' + a.FromSeq + '_' + d.InventoryNo IN @LnDeatilInfo";
                                    dynamicParameters.Add("LnDeatilInfo", changeLnDetailList.Select(x => x).ToArray());
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
                        data = "已更新資料【" + mainAffected + "】筆單頭, 【" + detailAffected + "】筆單身, 刪除【" + mainDelAffected + "】筆單頭, 【" + detailDelAffected + "】筆單身"
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

        #region //Delete
        #region //DeleteLotNumber -- 刪除批號資料 -- Ann 2024-03-22
        public string DeleteLotNumber(int LotNumberId)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷批號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.LotNumber
                                WHERE LotNumberId = @LotNumberId";
                        dynamicParameters.Add("LotNumberId", LotNumberId);

                        var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);
                        if (LotNumberResult.Count() <= 0) throw new SystemException("批號資料錯誤!");
                        #endregion

                        #region //確認此批號無任何單身交易紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM SCM.LnDetail a 
                                WHERE a.LotNumberId = @LotNumberId";
                        dynamicParameters.Add("LotNumberId", LotNumberId);

                        var LnDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (LnDetailResult.Count() > 0) throw new SystemException("此批號已經有交易紀錄，無法刪除!!");
                        #endregion

                        #region //確認此批號尚未綁定任何條碼
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM SCM.LnBarcode a 
                                WHERE a.LotNumberId = @LotNumberId";
                        dynamicParameters.Add("LotNumberId", LotNumberId);

                        var LnBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (LnDetailResult.Count() > 0) throw new SystemException("此批號已經有交易紀錄，無法刪除!!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.LotNumber
                                WHERE LotNumberId = @LotNumberId";
                        dynamicParameters.Add("LotNumberId", LotNumberId);

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

        #region //Api

        #endregion
    }
}
