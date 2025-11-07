using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.IO;
using System.Web;
using System.Linq;
using System.Transactions;
using System.Configuration;
using System.Data.SqlClient;
using System.Collections.Generic;
using SCMDA;

namespace QMSDA
{
    public class QcDataManagementDA
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
        public MamoHelper mamoHelper = new MamoHelper();

        public QcDataManagementDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            AutomationConnectionStrings = ConfigurationManager.AppSettings["AutomationDB"];

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
        #region //GetQcRecord -- 取得量測記錄資料 -- Ann 2023-02-24
        public string GetQcRecord(int QcRecordId, string WoErpFullNo, string QcNoticeNo, string MtlItemNo, string MtlItemName
            , int QcTypeId, string CheckQcMeasureData, int UserId, string StartDate, string EndDate, int MoId, int ModeId, string QcType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcRecordId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.QcNoticeId, a.QcTypeId, a.MoId, a.MoProcessId, ISNULL(a.PassQty, 0) PassQty, ISNULL(a.NgQty, 0) NgQty, a.SystemStatus
                        , a.QcStatus, a.QcUserId, ISNULL(a.Remark, '') Remark, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate, a.SpreadsheetData
                        , a.InputType, a.DefaultFileId, a.CheckQcMeasureData, a.CurrentFileId FileId, a.DefaultSpreadsheetData, a.SupportAqFlag QcSupportAqFlag
                        , a.SupportProcessFlag QcSupportProcessFlag, a.ResolveFileJson, a.UrgentFlag, a.MeasurementPlanning
                        , b.WoSeq, b.ModeId
                        , (c.WoErpPrefix + '-' + c.WoErpNo) WoErpFullNo, c.CompanyId
                        , d.MtlItemNo, d.MtlItemName
                        , ISNULL(e.ProcessAlias, '') ProcessAlias
                        , ISNULL(f.StatusName, '') StatusName
                        , g.UserNo, g.UserName
                        , h.QcNoticeNo, h.QcNoticeQty
                        , i.StatusName CheckQcMeasureDataName
                        , (
                            SELECT (
                                SELECT xa.ControlId, xa.Cad2DFile, xa.Pdf2DFile
                                , xb.CustomerMtlItemNo, xb.CustomerCadControlId
                                , xc.CadFile
                                , xd.[FileName] + xd.FileExtension Cad2DFileFullName
                                , xe.[FileName] + xe.FileExtension CustomerCadFileFullName
                                , xf.[FileName] + xf.FileExtension Pdf2DFileFullName
                                FROM MES.RoutingItem x
                                INNER JOIN PDM.RdDesignControl xa ON x.ControlId = xa.ControlId
                                INNER JOIN PDM.RdDesign xb ON xa.DesignId = xb.DesignId
                                INNER JOIN PDM.CustomerCadControl xc ON xb.CustomerCadControlId = xc.ControlId
                                INNER JOIN BAS.[File] xd ON xa.Cad2DFile = xd.FileId
                                INNER JOIN BAS.[File] xe ON xc.CadFile = xe.FileId
                                LEFT JOIN BAS.[File] xf ON xa.Pdf2DFile = xf.FileId
                                WHERE x.RoutingItemId = j.RoutingItemId
                                FOR JSON PATH, ROOT('data')
                            ) CadInfoData
                            FROM MES.MoRouting j
                            WHERE j.MoId = a.MoId
                            FOR JSON PATH, ROOT('data')
                        ) CadInfo
                        , (
                            SELECT (
                                SELECT xa.Cad2DFileAbsolutePath, xa.Pdf2DFileAbsolutePath
                                , xb.CustomerMtlItemNo, xb.CustomerCadControlId
                                , xc.CadFileAbsolutePath
                                FROM MES.RoutingItem x
                                INNER JOIN PDM.RdDesignControl xa ON x.ControlId = xa.ControlId
                                INNER JOIN PDM.RdDesign xb ON xa.DesignId = xb.DesignId
                                INNER JOIN PDM.CustomerCadControl xc ON xb.CustomerCadControlId = xc.ControlId
                                WHERE x.RoutingItemId = j.RoutingItemId
                                FOR JSON PATH, ROOT('data')
                            ) CadInfoData
                            FROM MES.MoRouting j
                            WHERE j.MoId = a.MoId
                            FOR JSON PATH, ROOT('data')
                        ) CadFileAbsolutePathInfo
                        , k.QcTypeNo, k.QcTypeName, k.SupportAqFlag, k.SupportProcessFlag
                        , (
                              SELECT x.FileId
                              , xa.[FileName], xa.FileExtension
                              FROM MES.QcRecordFile x
                              INNER JOIN BAS.[File] xa ON x.FileId = xa.FileId
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'other'
                              AND x.FileId IS NOT NULL
                              FOR JSON PATH, ROOT('data')
                        ) QcRecordFile
                        , (
                              SELECT x.QcRecordFileId, x.PhysicalPath
                              FROM MES.QcRecordFile x
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'other'
                              AND x.PhysicalPath IS NOT NULL
                              FOR JSON PATH, ROOT('data')
                        ) QcRecordFileByNas
                        , (
                              SELECT x.FileId
                              , xa.[FileName] + xa.FileExtension FileFullName
                              FROM MES.QcRecordFile x
                              INNER JOIN BAS.[File] xa ON x.FileId = xa.FileId
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'resolve'
                              FOR JSON PATH, ROOT('data')
                        ) ResolveFile";
                    sqlQuery.mainTables =
                        @"FROM MES.QcRecord a
                        INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                        INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                        INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                        LEFT JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                        LEFT JOIN BAS.[Status] f ON a.QcStatus = f.StatusNo AND f.StatusSchema = 'QcRecord.QcStatus'
                        INNER JOIN BAS.[User] g ON a.CreateBy = g.UserId
                        LEFT JOIN QMS.QcNotice h ON a.QcNoticeId = h.QcNoticeId
                        INNER JOIN QMS.QcType k ON a.QcTypeId = k.QcTypeId
                        INNER JOIN BAS.Status i ON a.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcRecordId", @" AND a.QcRecordId = @QcRecordId", QcRecordId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND (c.WoErpPrefix + '-' + c.WoErpNo +  '(' + CONVERT(VARCHAR(10), b.WoSeq) + ')') LIKE '%' + @WoErpFullNo + '%'", WoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeNo", @" AND h.QcNoticeNo LIKE '%' + @QcNoticeNo + '%'", QcNoticeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND d.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND d.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcTypeId", @" AND a.QcTypeId = @QcTypeId", QcTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CheckQcMeasureData", @" AND a.CheckQcMeasureData = @CheckQcMeasureData", CheckQcMeasureData);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.CreateBy = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND b.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcType", @" AND k.QcTypeNo = @QcType", QcType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "CASE WHEN a.UrgentFlag = 'Y' THEN 0 WHEN a.CheckQcMeasureData = 'S' THEN 1 WHEN a.CheckQcMeasureData = 'C' THEN 2 WHEN a.CheckQcMeasureData = 'A' THEN 3 WHEN a.CheckQcMeasureData = 'N' THEN 4 WHEN a.CheckQcMeasureData = 'P' THEN 5 WHEN a.CheckQcMeasureData = 'Y' THEN 6 ELSE 7 END, a.QcRecordId DESC";
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

        #region //GetLetteringBarcodeData -- 取得刻字條碼資料 -- Ann 2023-04-03
        public string GetLetteringBarcodeData(int MoId, int ItemSeq)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.BarcodeId, a.BarcodeNo, a.MoId
                            , b.ItemNo, b.ItemValue, b.ItemSeq
                            FROM MES.Barcode a
                            INNER JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering'
                            WHERE a.MoId = @MoId";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItemSeq", @" AND b.ItemSeq = @ItemSeq", ItemSeq);
                    dynamicParameters.Add("ItemSeq", ItemSeq);
                    dynamicParameters.Add("MoId", MoId);

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

        #region //GetBarcodeAccurately -- 驗證條碼是否可以進行量測數據上傳 -- Ann 2023-04-09
        public string GetBarcodeAccurately(int QcRecordId, string BarcodeListData, string QcType, int MoProcessId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認紀錄單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT CheckQcMeasureData
                            FROM MES.QcRecord a
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                    if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄單資料錯誤!!");

                    string CheckQcMeasureData = "";
                    foreach (var item in QcRecordResult)
                    {
                        if (item.CheckQcMeasureData == "Y") throw new SystemException("此紀錄已上傳!!");
                        CheckQcMeasureData = item.CheckQcMeasureData;
                    }
                    #endregion

                    #region //確認是否已上傳過量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                            FROM QMS.QcMeasureData a
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);

                    if (QcMeasureDataResult.Count() > 0 && CheckQcMeasureData != "P") throw new SystemException("此量測記錄單已上傳過量測記錄，無法重複上傳!!");
                    #endregion

                    var BarcodeList = BarcodeListData.Split(',');
                    foreach (var barcodeNo in BarcodeList)
                    {
                        #region //確認條碼資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.BarcodeId, a.CurrentProdStatus
                                , b.ItemValue
                                FROM MES.Barcode a
                                LEFT JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering'
                                WHERE a.BarcodeNo = @BarcodeNo";
                        dynamicParameters.Add("BarcodeNo", barcodeNo);

                        var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (BarcodeResult.Count() <= 0) throw new SystemException("條碼【" + barcodeNo + "】查無任何資料!!");

                        int BarcodeId = -1;
                        string ItemValue = "";
                        foreach (var item in BarcodeResult)
                        {
                            if (item.CurrentProdStatus != "P" && item.CurrentProdStatus != "M" && CheckQcMeasureData != "P") throw new SystemException("條碼【" + barcodeNo + "(" + item.ItemValue + ")】目前狀態非良品，不可進行量測上傳!!");
                            BarcodeId = item.BarcodeId;
                            ItemValue = item.ItemValue;
                        }
                        #endregion

                        #region //確認條碼是否完工
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 c.CurrentProdStatus
                                FROM MES.BarcodeProcess a
                                INNER JOIN MES.Barcode c ON a.BarcodeId = c.BarcodeId
                                WHERE a.BarcodeId = @BarcodeId
                                AND a.FinishDate IS NULL";
                        dynamicParameters.Add("BarcodeId", BarcodeId);

                        var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in BarcodeProcessResult)
                        {
                            if (CheckQcMeasureData != "P") throw new SystemException("條碼【" + barcodeNo + "(" + ItemValue + ")】目前為加工狀態，無法進行量測數據上傳!!");
                            if (item.CurrentProdStatus == "F") throw new SystemException("條碼【" + barcodeNo + "(" + ItemValue + ")】目前為加工狀態，無法進行量測數據上傳!!");
                        }
                        #endregion

                        #region //若為工程檢，檢核條碼是否正在此製程，且已完成加工流程
                        if (QcType == "IPQC")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.CurrentMoProcessId
                                    , b.ProcessAlias
                                    FROM MES.Barcode a
                                    INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
                                    WHERE a.BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", barcodeNo);

                            var CheckIPQCBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in CheckIPQCBarcodeResult)
                            {
                                if (item.CurrentMoProcessId != MoProcessId) throw new SystemException("條碼【" + barcodeNo + ")】目前為【" + item.ProcessAlias + "】完工，無法進行此製程工程檢!!");
                            }
                        }
                        #endregion

                        #region //確認條碼是否存在品異單且未完成判定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.AqBarcode a
                                WHERE a.BarcodeId = @BarcodeId
                                AND a.JudgeStatus IS NULL";
                        dynamicParameters.Add("BarcodeId", BarcodeId);

                        var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (AqBarcodeResult.Count() > 0) throw new SystemException("條碼【" + barcodeNo + "(" + ItemValue + ")】目前尚為品異單綁定條碼，且未完成判定!!");
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = "EMPTY"
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

        #region //GetLotNumberAccurately -- 驗證條碼是否可以進行量測數據上傳 -- Ann 2023-04-09
        public string GetLotNumberAccurately(int QcRecordId, string LotNumberListData)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認紀錄單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CheckQcMeasureData
                            , b.GrDetailId
                            , c.MtlItemId
                            FROM MES.QcRecord a
                            INNER JOIN MES.QcGoodsReceipt b ON a.QcRecordId = b.QcRecordId 
                            INNER JOIN SCM.GrDetail c ON b.GrDetailId = c.GrDetailId
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                    if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄單資料錯誤!!");

                    string CheckQcMeasureData = "";
                    int MtlItemId = -1;
                    int GrDetailId = -1;
                    foreach (var item in QcRecordResult)
                    {
                        if (item.CheckQcMeasureData == "Y") throw new SystemException("此紀錄已上傳!!");
                        CheckQcMeasureData = item.CheckQcMeasureData;
                        MtlItemId = item.MtlItemId;
                        GrDetailId = item.GrDetailId;
                    }
                    #endregion

                    #region //確認是否已上傳過量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                            FROM QMS.QcMeasureData a
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);

                    if (QcMeasureDataResult.Count() > 0 && CheckQcMeasureData != "P") throw new SystemException("此量測記錄單已上傳過量測記錄，無法重複上傳!!");
                    #endregion

                    var LotNumberList = LotNumberListData.Split(',');
                    foreach (var lotNumber in LotNumberList)
                    {
                        #region //確認批號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.GrDetail a 
                                WHERE a.GrDetailId = @GrDetailId
                                AND a.LotNumber = @LotNumber";
                        dynamicParameters.Add("GrDetailId", GrDetailId);
                        dynamicParameters.Add("LotNumber", lotNumber);

                        var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                        if (LotNumberResult.Count() <= 0) throw new SystemException("查無此批號【" + lotNumber + "】資料!!");
                        #endregion

                        #region //確認此進貨單身是否有品異單未判定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.JudgeStatus
                                , b.AbnormalqualityNo
                                FROM QMS.AqBarcode a 
                                INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                WHERE b.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in AqBarcodeResult)
                        {
                            if (item.JudgeStatus == null) throw new SystemException("此進貨單單身尚有品異單【" + item.AbnormalqualityNo + "】未判定!!");
                        }
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = "EMPTY"
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

        #region //GetQmmDetail -- 取得量測機台資料 -- Ann 2023-04-12
        public string GetQmmDetail(int QmmDetailId, int ShopId, string MachineNumber, int QcMachineModeId, string QcMachineModeList)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QmmDetailId, a.QcMachineModeId, a.MachineNumber, a.MachineId
                            , b.MachineDesc, b.MachineDesc MachineName
                            , d.QcMachineModeName
                            , c.ShopName + '-' + b.MachineDesc FullMachineName
                            FROM QMS.QmmDetail a
                            INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                            INNER JOIN MES.WorkShop c ON b.ShopId = c.ShopId
                            INNER JOIN QMS.QcMachineMode d ON a.QcMachineModeId = d.QcMachineModeId
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineNumber", @" AND a.MachineNumber = @MachineNumber", MachineNumber);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ShopId", @" AND c.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcMachineModeId", @" AND a.QcMachineModeId = @QcMachineModeId", QcMachineModeId);
                    if (QcMachineModeList.Length > 0)
                    {
                        sql += @" AND a.QcMachineModeId IN @QcMachineModeList";
                        dynamicParameters.Add("QcMachineModeList", QcMachineModeList.Split(','));
                    }
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcMachineModeList", @" AND a.QcMachineModeId IN @QcMachineModeList", QcMachineModeList.Split(','));

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

        #region //GetWorkShop -- 取得車間資料 -- Ann 2023-04-12
        public string GetWorkShop(int ShopId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ShopId, a.ShopNo, a.ShopName
                            FROM MES.WorkShop a
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);

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

        #region //GetBarcodeData -- 取得條碼資料 -- Ann 2023-04-16
        public string GetBarcodeData(int MoId, string BarcodeNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.BarcodeId, b.ItemValue
                            FROM MES.Barcode a
                            LEFT JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering'
                            WHERE a.BarcodeNo = @BarcodeNo
                            AND a.MoId = @MoId";
                    dynamicParameters.Add("BarcodeNo", BarcodeNo);
                    dynamicParameters.Add("MoId", MoId);

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

        #region //GetQcMeasureData -- 取得量測資料 -- Ann 2023-05-25
        public string GetQcMeasureData(int QcRecordId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcRecordId, a.QcItemId, a.QcItemDesc, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                            , a.ZAxis, a.BarcodeId, a.QmmDetailId, a.MeasureValue, a.QcResult, a.CauseId, a.CauseDesc NgCodeDesc
                            , a.Cavity, a.MakeCount, a.CellHeader, a.Row, a.Remark, a.LotNumber
                            , b.BarcodeNo
                            , c.QcItemNo ItemNo, c.QcItemName
                            , e.ItemValue LetterValue
                            , f.MachineNumber
                            , g.MachineDesc QcMachine
                            , h.CauseNo NgCode
                            , i.UserNo QcUserNo, i.UserName QcUserName
                            FROM QMS.QcMeasureData a
                            LEFT JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                            INNER JOIN QMS.QcItem c ON a.QcItemId = c.QcItemId
                            INNER JOIN MES.QcRecord d ON a.QcRecordId = d.QcRecordId
                            LEFT JOIN MES.BarcodeAttribute e ON a.BarcodeId = e.BarcodeId AND e.ItemNo = 'Lettering' AND d.MoId = e.MoId
                            INNER JOIN QMS.QmmDetail f ON a.QmmDetailId = f.QmmDetailId
                            INNER JOIN MES.Machine g ON f.MachineId = g.MachineId
                            LEFT JOIN QMS.DefectCause h ON a.CauseId = h.CauseId
                            LEFT JOIN BAS.[User] i ON a.QcUserId = i.UserId
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

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

        #region //GetCheckQcMeasureData -- 取得量測單據狀態 -- Ann 2023-06-08
        public string GetCheckQcMeasureData(int QcRecordId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CheckQcMeasureData, a.SupportAqFlag, a.SupportProcessFlag
                            FROM MES.QcRecord a 
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

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

        #region //GetQcItemPrinciple -- 取得量測項目編碼原則 -- Ann 2023-10-03        
        public string GetQcItemPrinciple(int PrincipleId, int QcClassId, string PrincipleNo, string PrincipleDesc, int QmmDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PrincipleId, a.PrincipleNo, a.PrincipleDesc
                            , b.QcClassNo + a.PrincipleNo PrincipleFullNo
                            , ISNULL(d.QcItemId, -1) QcItemId, ISNULL(d.QcItemName, '') QcItemName, ISNULL(d.QcItemDesc, '') QcItemDesc
                            FROM QMS.QcItemPrinciple a 
                            INNER JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                            INNER JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                            LEFT JOIN QMS.QcItem d ON d.QcItemNo = b.QcClassNo + a.PrincipleNo AND d.QcClassId = a.QcClassId
                            INNER JOIN QMS.QmmDetail e ON a.QmmDetailId = e.QmmDetailId
                            WHERE c.CompanyId = @CurrentCompany";
                    if (CurrentCompany == 0) CurrentCompany = 2;
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleNo", @" AND a.PrincipleNo = @PrincipleNo", PrincipleNo);
                    if (PrincipleDesc.Length > 0)
                    {
                        sql += @" AND a.PrincipleDesc = N'" + PrincipleDesc + "'";
                    }
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleDesc", @" AND a.PrincipleDesc = 'N' + @PrincipleDesc", PrincipleDesc);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassId", @" AND a.QcClassId = @QcClassId", QcClassId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrincipleDesc
                                , b.PrincipleId, b.PrincipleNo
                                , c.QcClassNo + b.PrincipleNo PrincipleFullNo
                                , ISNULL(e.QcItemId, -1) QcItemId, ISNULL(e.QcItemName, '') QcItemName, ISNULL(e.QcItemDesc, '') QcItemDesc
                                FROM QMS.PrincipleDetail a 
                                INNER JOIN QMS.QcItemPrinciple b ON a.PrincipleId = b.PrincipleId
                                INNER JOIN QMS.QcClass c ON b.QcClassId = c.QcClassId
                                INNER JOIN QMS.QcGroup d ON c.QcGroupId = d.QcGroupId
                                LEFT JOIN QMS.QcItem e ON e.QcItemNo = c.QcClassNo + b.PrincipleNo AND e.QcClassId = b.QcClassId
                                INNER JOIN QMS.QmmDetail f ON b.QmmDetailId = f.QmmDetailId
                                WHERE d.CompanyId = @CompanyId";
                        if (CurrentCompany == 0) CurrentCompany = 2;
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleNo", @" AND b.PrincipleNo = @PrincipleNo", PrincipleNo);
                        //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleDesc", @" AND a.PrincipleDesc = 'N' + @PrincipleDesc", PrincipleDesc);
                        if (PrincipleDesc.Length > 0)
                        {
                            sql += @" AND a.PrincipleDesc = N'" + PrincipleDesc + "'";
                        }
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassId", @" AND b.QcClassId = @QcClassId", QcClassId);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND b.QmmDetailId = @QmmDetailId", QmmDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                    }

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

        #region //GetQcType -- 取得量測類型資料 -- Ann 2023-10-11
        public string GetQcType(int QcTypeId, int ModeId, string QcTypeNo, string QcTypeName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcTypeId, a.ModeId, a.QcTypeNo, a.QcTypeName
                            , a.QcTypeNo + ' ' + a.QcTypeName 'QcTypeWithNo'
                            , a.SupportAqFlag, a.SupportProcessFlag, a.[Status]
                            FROM QMS.QcType a 
                            INNER JOIN MES.ProdMode b ON a.ModeId = b.ModeId
                            WHERE b.CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcTypeId", @" AND a.QcTypeId = @QcTypeId", QcTypeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcTypeNo", @" AND a.QcTypeNo LIKE '%' + @QcTypeNo + '%'", QcTypeNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcTypeName", @" AND a.QcTypeName LIKE '%' + @QcTypeName + '%'", QcTypeName);

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

        #region //GetQcpRelationship -- 取得量測項目關聯別名資料 -- Ann 2023-10-18
        public string GetQcpRelationship(int QcprId, int PrincipleId, int QmmDetailId, string RelationshipItem)
        {
            if (QmmDetailId <= 0) throw new SystemException("量測機台不能為空!!");
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcprId, a.PrincipleId, a.QmmDetailId, a.RelationshipItem
                            , c.QcClassNo + b.PrincipleNo PrincipleFullNo
                            , ISNULL(e.QcItemId, -1) QcItemId, e.QcItemName, e.QcItemDesc
                            FROM QMS.QcpRelationship a 
                            INNER JOIN QMS.QcItemPrinciple b ON a.PrincipleId = b.PrincipleId
                            INNER JOIN QMS.QcClass c ON b.QcClassId = c.QcClassId
                            INNER JOIN QMS.QcGroup d ON c.QcGroupId = d.QcGroupId
                            LEFT JOIN QMS.QcItem e ON e.QcItemNo = c.QcClassNo + b.PrincipleNo
                            WHERE d.CompanyId = @CompanyId
                            AND a.QmmDetailId = @QmmDetailId
                            AND a.RelationshipItem = N'" + RelationshipItem + "'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("QmmDetailId", QmmDetailId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcprId", @" AND a.QcprId = @QcprId", QcprId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RelationshipItem", @" AND a.RelationshipItem = N@RelationshipItem", RelationshipItem);

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

        #region //GetFileInfo -- 取得檔案資料 -- Ann 2023-11-08
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
                    result = result.OrderBy(a => a.FileName);
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

        #region //GetQcFileParsing -- 取得檔案解析方式 -- Ann 2023-11-10
        public string GetQcFileParsing(int QfpId, int QmmDetailId, string FunctionName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QfpId, a.QmmDetailId, a.FunctionName, a.Encoding
                            FROM QMS.QcFileParsing a
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QfpId", @" AND a.QfpId = @QfpId", QfpId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FunctionName", @" AND a.FunctionName = @FunctionName", FunctionName);

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

        #region //GetQcMeasureDataHistory -- 取得量測資料歷程記錄 -- Ann 2023-11-13
        public string GetQcMeasureDataHistory(string MtlItemNo, string QcItemNo)
        {
            try
            {
                if (MtlItemNo.Length <= 0) throw new SystemException("【品號】不能為空!");
                if (QcItemNo.Length <= 0) throw new SystemException("【量測項目】不能為空!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DesignValue, a.UpperTolerance, a.LowerTolerance
                            FROM QMS.QcMeasureData a 
                            INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                            INNER JOIN MES.ManufactureOrder c ON b.MoId = c.MoId
                            INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                            INNER JOIN PDM.MtlItem e ON  d.MtlItemId = e.MtlItemId
                            INNER JOIN QMS.QcItem f ON a.QcItemId = f.QcItemId
                            WHERE e.MtlItemNo = @MtlItemNo
                            AND f.QcItemNo = @QcItemNo";
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                    dynamicParameters.Add("QcItemNo", QcItemNo);

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

        #region //GetQcClass -- 取得量測類型 -- Ann 2023-11-15
        public string GetQcClass(int QcClassId, string QcClassNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcClassId, a.QcGroupId, a.QcClassNo, a.QcClassName, a.QcClassDesc
                            , a.QcClassNo + ' ' + a.QcClassName QcClassFullNo
                            FROM QMS.QcClass a 
                            INNER JOIN QMS.QcGroup b ON a.QcGroupId = b.QcGroupId
                            WHERE b.CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassId", @" AND a.QcClassId = @QcClassId", QcClassId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassNo", @" AND a.QcClassNo = @QcClassNo", QcClassNo);

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

        #region //GetQcRecordFile -- 取得量測檔案資料 -- Ann 2023-11-08
        public string GetQcRecordFile(int QcRecordId, string FileIdList, string FileType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcRecordFileId, a.QcRecordId, a.FileType, a.FileId, a.BarcodeId, a.PhysicalPath
                            , a.InputType, a.Cavity, ISNULL(a.LotNumber, '') LotNumber, a.EffectiveDiameter
                            , ISNULL(b.BarcodeNo, '') BarcodeNo
                            , c.ItemValue LetteringNo, c.ItemSeq
                            , d.[FileName] + d.FileExtension FileFullName
                            , e.MoId
                            FROM MES.QcRecordFile a 
                            LEFT JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                            LEFT JOIN MES.BarcodeAttribute c ON c.BarcodeId = b.BarcodeId AND c.ItemNo = 'Lettering'
                            LEFT JOIN BAS.[File] d ON a.FileId = d.FileId
                            INNER JOIN MES.QcRecord e ON a.QcRecordId = e.QcRecordId
                            WHERE 1=1";

                    if (FileIdList.Length > 0)
                    {
                        sql += @" AND a.FileId IN (" + FileIdList + ")";
                    }

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcRecordId", @" AND a.QcRecordId = @QcRecordId", QcRecordId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FileType", @" AND a.FileType = @FileType", FileType);

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

        #region //GetBarcodeInfo -- 取得條碼相關資料 -- Ann 2024-01-04
        public string GetBarcodeInfo(string BarcodeNo, string InputType, int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.BarcodeId, a.BarcodeNo
                            , ISNULL(b.ItemValue, '') LetteringNo
                            FROM MES.Barcode a 
                            LEFT JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering' AND b.MoId = a.MoId
                            WHERE 1=1";

                    if (InputType == "BarcodeNo")
                    {
                        sql += @" AND a.BarcodeNo = @BarcodeNo";
                    }
                    else if (InputType == "Lettering")
                    {
                        sql += @" AND b.MoId = @MoId AND b.ItemSeq = @BarcodeNo";
                    }
                    else
                    {
                        throw new SystemException("此輸入方式【" + InputType + "】目前不支援!!");
                    }

                    dynamicParameters.Add("MoId", MoId);
                    dynamicParameters.Add("BarcodeNo", BarcodeNo);

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

        #region //GetQcGoodsReceipt -- 取得進貨檢單據資料 -- Ann 2024-04-01 
        public string GetQcGoodsReceipt(int QcGoodsReceiptId, int QcRecordId, string GrErpFullNo, string MtlItemNo, string MtlItemName
            , string CheckQcMeasureData, int UserId, string StartDate, string EndDate, string QcType
            , string PoErpFullNo, int PoUserId, string PrErpFullNo, int PrUserId

            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcGoodsReceiptId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.QcRecordId, a.GrDetailId, a.QcStatus, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                        , b.QcRecordId, b.Remark, b.CheckQcMeasureData, b.InputType, b.ResolveFileJson, b.DefaultSpreadsheetData, b.SpreadsheetData
                        , c.GrDetailId, c.GrSeq, c.ReceiptQty, c.LotNumber, c.ConfirmStatus
                        , d.GrId, d.GrErpPrefix, d.GrErpNo, d.CompanyId
                        , e.MtlItemNo, e.MtlItemName, e.MtlItemSpec
                        , f.UserNo, f.UserName
                        , g.StatusName QcStatusName
                        , h.SupplierNo, h.SupplierShortName
                        , i.StatusName CheckQcMeasureDataName
                        , (
                              SELECT x.FileId
                              , xa.[FileName], xa.FileExtension
                              FROM MES.QcRecordFile x
                              INNER JOIN BAS.[File] xa ON x.FileId = xa.FileId
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'other'
                              FOR JSON PATH, ROOT('data')
                        ) QcRecordFile
                        , (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq) PoErpFullNo, (xd.UserNo + '-' + xd.UserName) PoConfidmUser
						, (xf.PrErpPrefix + '-' + xf.PrErpNo + '-' + xe.PrSequence) PrErpFullNo, (xg.UserNo + '-' + xg.UserName) PrConfidmUser";
                    sqlQuery.mainTables =
                        @"FROM MES.QcGoodsReceipt a 
                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                        INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                        INNER JOIN SCM.GoodsReceipt d ON c.GrId = d.GrId
                        INNER JOIN PDM.MtlItem e ON c.MtlItemId = e.MtlItemId
                        INNER JOIN BAS.[User] f ON a.CreateBy = f.UserId
                        INNER JOIN BAS.[Status] g ON a.QcStatus = g.StatusNo AND g.StatusSchema = 'QcGoodsReceipt.QcStatus'
                        INNER JOIN SCM.Supplier h ON d.SupplierId = h.SupplierId
                        INNER JOIN BAS.Status i ON b.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'
                        INNER JOIN SCM.PoDetail xb ON (c.PoErpPrefix + '-' + c.PoErpNo + '-' + c.PoSeq) =  (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
						INNER JOIN SCM.PurchaseOrder xc ON xc.PoId = xb.PoId
						left JOIN BAS.[User] xd ON xc.PoUserId = xd.UserId
						left JOIN SCM.PurchaseRequisition xf ON (xb.PrErpPrefix + '-' + xb.PrErpNo) =  (xf.PrErpPrefix + '-' + xf.PrErpNo)
						left JOIN SCM.PrDetail xe ON xe.PrId = xf.PrId AND xb.PrSequence = xe.PrSequence
                        left JOIN BAS.[User] xg ON xf.UserId = xg.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGoodsReceiptId", @" AND a.QcGoodsReceiptId = @QcGoodsReceiptId", QcGoodsReceiptId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcRecordId", @" AND a.QcRecordId = @QcRecordId", QcRecordId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrErpFullNo", @" AND d.GrErpPrefix + '-' + d.GrErpNo LIKE '%' + @GrErpFullNo + '%'", GrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND e.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CheckQcMeasureData", @" AND b.CheckQcMeasureData = @CheckQcMeasureData", CheckQcMeasureData);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.CreateBy = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpFullNo", @" AND xb.PoErpPrefix + '-' + xb.PoErpNo LIKE '%' + @PoErpFullNo + '%'", PoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoUserId", @" AND xc.PoUserId = @PoUserId", PoUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpFullNo", @" AND xf.PrErpPrefix + '-' + xf.PrErpNo  LIKE '%' + @PrErpFullNo + '%'", PrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrUserId", @" AND xf.UserId = @PrUserId", PrUserId);




                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcGoodsReceiptId DESC";
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

        #region //GetQcRecordPlanning -- 取得量測單據排程資料 -- Ann 2024-08-20
        public string GetQcRecordPlanning(int QrpId, int QcRecordId, int QcMachineModeId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QrpId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.QcRecordId, a.QcMachineModeId, a.TotalQcItemCount, a.EstimatedMeasurementTime, a.ConfirmStatus
                        , b.CheckQcMeasureData
                        , c.MoId
                        , e.MtlItemId, e.MtlItemNo, e.MtlItemName, e.MtlItemSpec
                        , f.QcMachineModeNo, f.QcMachineModeName";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcRecordPlanning a 
                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                        INNER JOIN MES.ManufactureOrder c ON b.MoId = c.MoId
                        INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                        INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                        INNER JOIN QMS.QcMachineMode f ON a.QcMachineModeId = f.QcMachineModeId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QrpId", @" AND a.QrpId = @QrpId", QrpId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcRecordId", @" AND a.QcRecordId = @QcRecordId", QcRecordId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeId", @" AND a.QcMachineModeId = @QcMachineModeId", QcMachineModeId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QrpId DESC";
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

        #region //GetQcRecordQcItem -- 取得量測單據排程資料 -- Ann 2024-08-20
        public string GetQcRecordQcItem(int QcRecordId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT SpreadsheetData, PlanningSpreadsheetDate
                            FROM MES.QcRecord
                            WHERE QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

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

        #region //GetQcItem -- 取得量測項目 -- Ann 2024-08-23
        public string GetQcItem(int QcItemId, string QcItemNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcItemId, a.QcItemNo, a.QcItemName, a.QcItemDesc
                            FROM QMS.QcItem a 
                            INNER JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                            INNER JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId", QcItemId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemNo", @" AND a.QcItemNo = @QcItemNo", QcItemNo);

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

        #region //GetQcMachineMode -- 取得量測機型、機台資料 -- Ann 2024-08-23
        public string GetQcMachineMode(int QmmDetailId, string MachineNumber)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QmmDetailId, a.MachineNumber
                            , b.QcMachineModeId, b.QcMachineModeNo, b.QcMachineModeName
                            FROM QMS.QmmDetail a 
                            INNER JOIN QMS.QcMachineMode b ON a.QcMachineModeId = b.QcMachineModeId
                            WHERE b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineNumber", @" AND a.MachineNumber = @MachineNumber", MachineNumber);

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

        #region //GetQcMachineModeInfo -- 取得量測機型資料 -- Ann 2024-09-18
        public string GetQcMachineModeInfo(int QcMachineModeId, string QcMachineModeNumber, string QcMachineModeNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcMachineModeId, a.QcMachineModeNumber, a.QcMachineModeNo
                            , a.QcMachineModeName, a.QcMachineModeDesc
                            FROM QMS.QcMachineMode a 
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcMachineModeId", @" AND a.QcMachineModeId = @QcMachineModeId", QcMachineModeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcMachineModeNumber", @" AND a.QcMachineModeNumber = @QcMachineModeNumber", QcMachineModeNumber);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcMachineModeNo", @" AND a.QcMachineModeNo = @QcMachineModeNo", QcMachineModeNo);

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

        #region //GetAllQcMachineMode -- 取得量測機型資料 -- Ann 2024-08-26
        public string GetAllQcMachineMode()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcMachineModeId, a.QcMachineModeNo, a.QcMachineModeName
                            , CAST(a.QcMachineModeId AS nvarchar) + ':' + a.QcMachineModeName QcMachineModeFullNo
                            FROM QMS.QcMachineMode a 
                            WHERE a.CompanyId = @CompanyId";
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

        #region //GetQcMachineModePlanning -- 取得量測機台排程資料 -- Ann 2024-08-27
        public string GetQcMachineModePlanning(int QmmpId, int QrpId, int QmmDetailId, string StartDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QmmpId, a.QrpId, a.QmmDetailId
                            , FORMAT(a.EstimatedStartDate, 'MM-dd HH:mm:ss') EstimatedStartDate
                            , FORMAT(a.EstimatedEndDate, 'MM-dd HH:mm:ss') EstimatedEndDate
                            , FORMAT(a.StartDate, 'MM-dd HH:mm:ss') StartDate
                            , FORMAT(a.EndDate, 'MM-dd HH:mm:ss') EndDate
                            , a.Sort
                            , d.QcRecordId
                            , h.MtlItemId, h.MtlItemNo, h.MtlItemName
                            FROM QMS.QcMachineModePlanning a 
                            INNER JOIN QMS.QmmDetail b ON a.QmmDetailId = b.QmmDetailId
                            INNER JOIN QMS.QcMachineMode c ON b.QcMachineModeId = c.QcMachineModeId
                            INNER JOIN QMS.QcRecordPlanning d ON a.QrpId = d.QrpId
                            INNER JOIN MES.QcRecord e ON d.QcRecordId = e.QcRecordId
                            INNER JOIN MES.ManufactureOrder f ON e.MoId = f.MoId
                            INNER JOIN MES.WipOrder g ON f.WoId = g.WoId
                            INNER JOIN PDM.MtlItem h ON g.MtlItemId = h.MtlItemId
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmpId", @" AND a.QmmpId = @QmmpId", QmmpId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QrpId", @" AND a.QrpId = @QrpId", QrpId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @StartDate", StartDate);

                    sql += @" ORDER BY a.QmmDetailId, a.Sort";

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

        #region //GetQcMachineModePlanningForKanban -- 取得量測機台排程資料(For量測看板) -- Ann 2024-09-30
        public string GetQcMachineModePlanningForKanban(string QcMachineModeList, string StartDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QmmpId, a.QrpId, a.QmmDetailId
                            , ISNULL(FORMAT(a.EstimatedStartDate, 'MM-dd HH:mm:ss'), '') EstimatedStartDate
                            , ISNULL(FORMAT(a.EstimatedEndDate, 'MM-dd HH:mm:ss'), '') EstimatedEndDate
                            , ISNULL(FORMAT(a.StartDate, 'MM-dd HH:mm:ss'), '') StartDate
                            , ISNULL(FORMAT(a.EndDate, 'MM-dd HH:mm:ss'), '') EndDate
                            , a.Sort, a.Status
                            , d.QcRecordId
                            , h.MtlItemId, h.MtlItemNo, h.MtlItemName
                            , i.MachineNo, i.MachineDesc
                            , j.StatusName
                            , DATEDIFF(HOUR, ISNULL(a.StartDate, GETDATE()), a.EstimatedStartDate) DateDifferenceInDays
                            FROM QMS.QcMachineModePlanning a 
                            INNER JOIN QMS.QmmDetail b ON a.QmmDetailId = b.QmmDetailId
                            INNER JOIN QMS.QcMachineMode c ON b.QcMachineModeId = c.QcMachineModeId
                            INNER JOIN QMS.QcRecordPlanning d ON a.QrpId = d.QrpId
                            INNER JOIN MES.QcRecord e ON d.QcRecordId = e.QcRecordId
                            INNER JOIN MES.ManufactureOrder f ON e.MoId = f.MoId
                            INNER JOIN MES.WipOrder g ON f.WoId = g.WoId
                            INNER JOIN PDM.MtlItem h ON g.MtlItemId = h.MtlItemId
                            INNER JOIN MES.Machine i ON b.MachineId = i.MachineId
                            INNER JOIN BAS.[Status] j ON j.StatusSchema = 'QcMachineModePlanning.Status' AND j.StatusNo = a.[Status]
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcMachineModeList", @" AND b.QcMachineModeId IN @QcMachineModeList", QcMachineModeList.Split(','));
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND FORMAT(a.EstimatedStartDate, 'yyyy-MM-dd') = @StartDate", StartDate);

                    sql += @" ORDER BY a.QmmDetailId, a.Sort";

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

        #region //GetQcMeasureDataTemp -- 取得量測暫存資料 -- Ann 2024-10-15
        public string GetQcMeasureDataTemp(int QcRecordId)
        {
            try
            {
                if (QcRecordId <= 0) throw new SystemException("量測單號不能為空!!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TempId, a.QcRecordId, a.Cell, a.Css, a.[Format], a.[Value]
                            FROM QMS.QcMeasureDataTemp a 
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    sql += @" ORDER BY a.TempId";

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

        #region //GetQcMachineModePlanningSort -- 取得量測機台排程最大排序 -- Ann 2024-11-06
        public string GetQcMachineModePlanningSort(int QmmDetailId)
        {
            try
            {
                if (QmmDetailId <= 0) throw new SystemException("量測排程機台不能為空!!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT COUNT(1) Count
                            FROM QMS.QcMachineModePlanning a
                            WHERE a.QmmDetailId = @QmmDetailId
                            AND CONVERT(DATE, a.CreateDate) = CONVERT(DATE, GETDATE())";
                    dynamicParameters.Add("QmmDetailId", QmmDetailId);

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

        #region //GetQcGoodsReceiptLog -- 取得進貨檢驗作業資料 -- Ann 2025-02-07
        public string GetQcGoodsReceiptLog(int LogId, int QcGoodsReceiptId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.LogId, a.QcGoodsReceiptId, a.ReceiptQty, a.AcceptQty, a.ReturnQty
                            , FORMAT(a.AcceptanceDate, 'yyyy-MM-dd') AcceptanceDate, a.QcStatus, a.QuickStatus, a.Remark
                            , c.GrSeq, c.InventoryId
                            , d.GrErpPrefix, d.GrErpNo
                            , e.MtlItemNo, e.MtlItemName, e.MtlItemSpec
                            , (
	                            SELECT x.FileId
	                            , xa.[FileName], xa.FileExtension
	                            FROM MES.QcGoodsReceiptLogFile x
	                            INNER JOIN BAS.[File] xa ON x.FileId = xa.FileId
	                            WHERE x.LogId = a.LogId
	                            FOR JSON PATH, ROOT('data')
                            ) QcGoodsReceiptLogFile
                            FROM MES.QcGoodsReceiptLog a 
                            INNER JOIN MES.QcGoodsReceipt b ON a.QcGoodsReceiptId = b.QcGoodsReceiptId
                            INNER JOIN SCM.GrDetail c ON b.GrDetailId = c.GrDetailId
                            INNER JOIN SCM.GoodsReceipt d ON c.GrId = d.GrId
                            INNER JOIN PDM.MtlItem e ON c.MtlItemId = e.MtlItemId
                            WHERE d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "LogId", @" AND a.LogId = @LogId", LogId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcGoodsReceiptId", @" AND a.QcGoodsReceiptId = @QcGoodsReceiptId", QcGoodsReceiptId);

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

        #region //GetLotNumber -- 取得批號資料 -- Ann 2024-08-29
        public string GetLotNumber(string MtlItemNo, string LotNumber)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.LotNumberId, a.LotNumberNo, a.MtlItemId
                            , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                            FROM SCM.LotNumber a 
                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                            WHERE b.MtlItemNo = @MtlItemNo
                            AND a.LotNumberNo = @LotNumberNo 
                            AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                    dynamicParameters.Add("LotNumberNo", LotNumber);
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

        #region //GetQcWhitelist -- 取得上傳路徑白名單 -- Ann 2024-09-23
        public string GetQcWhitelist(int ListId, int DepartmentId, string FolderPath)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得使用者部門
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DepartmentId
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", CurrentUser);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    if (UserResult.Count() <= 0) throw new SystemException("使用者資訊錯誤!!");

                    foreach (var item in UserResult)
                    {
                        DepartmentId = item.DepartmentId;
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ListId, a.DepartmentId, a.FolderPath
                            FROM QMS.QcWhitelist a 
                            WHERE a.DepartmentId = @DepartmentId";
                    dynamicParameters.Add("DepartmentId", DepartmentId);

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

        #region //GetFileInfoById -- 取得檔案資料 -- Ann 2025-04-29
        public BmFileInfo GetFileInfoById(int FileId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FileId, FileName, FileContent, FileExtension, FileSize, ClientIP, Source, DeleteStatus
                            FROM BAS.[File]
                            WHERE FileId = @FileId";
                    dynamicParameters.Add("FileId", FileId);

                    var result = sqlConnection.Query<BmFileInfo>(sql, dynamicParameters).FirstOrDefault();

                    return result;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return new BmFileInfo();
        }
        #endregion

        #region//GetTrayNoToBarcodeNo Tray轉條碼
        public string GetTrayNoToBarcodeNo(string Company, string InputNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", Company);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                    foreach (var item in companyResult)
                    {
                        CurrentCompany = item.CompanyId;
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT BarcodeNo
                            FROM MES.Tray
                            WHERE CompanyId = @CurrentCompany
                            AND TrayNo = @InputNo";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);
                    dynamicParameters.Add("InputNo", InputNo);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0)
                    {
                        sql = @"SELECT a.BarcodeNo
                            FROM MES.Barcode a
                            INNER JOIN MES.ManufactureOrder b ON a.MoId=b.MoId
                            INNER JOIN MES.WipOrder c ON c.WoId=b.WoId
                            INNER JOIN BAS.Company d ON d.CompanyId=c.CompanyId
                            WHERE a.CompanyId = @CurrentCompany
                            AND a.BarcodeNo = @InputNo";
                        dynamicParameters.Add("CurrentCompany", CurrentCompany);
                        dynamicParameters.Add("InputNo", InputNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                    }

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

        #region //GetQcRecordFileID -- 取得量測檔案路徑 -- WuTc 2024-12-16
        public string GetQcRecordFileID(int QcRecordFileId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcRecordFileId, a.QcRecordId, a.FileType, a.PhysicalPath, ISNULL(a.BarcodeId, 0) BarcodeId, a.LotNumber
                            , (SELECT CompanyId FROM BAS.[User] u 
	                            INNER JOIN BAS.Department d ON u.DepartmentId = d.DepartmentId
	                            WHERE u.UserId = a.CreateBy) CompanyId
                            FROM MES.QcRecordFile a 
                                WHERE a.QcRecordFileId = @QcRecordFileId";

                    dynamicParameters.Add("QcRecordFileId", QcRecordFileId);

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

        #region //GetOutsourcingQcRecord -- 取得量測記錄資料 -- GPAI 2025-04-10
        public string GetOutsourcingQcRecord(int QcRecordId, string QcNoticeNo, string MtlItemNo, string MtlItemName, int MtlItemId
            , int QcTypeId, string CheckQcMeasureData, int UserId, string StartDate, string EndDate, string QcType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcRecordId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.QcNoticeId, a.QcTypeId, a.MoId, a.MoProcessId, ISNULL(a.PassQty, 0) PassQty, ISNULL(a.NgQty, 0) NgQty, a.SystemStatus
                        , a.QcStatus, a.QcUserId, ISNULL(a.Remark, '') Remark, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate, a.SpreadsheetData
                        , a.InputType, a.DefaultFileId, a.CheckQcMeasureData, a.CurrentFileId FileId, a.DefaultSpreadsheetData, a.SupportAqFlag QcSupportAqFlag
                        , a.SupportProcessFlag QcSupportProcessFlag, a.ResolveFileJson, a.UrgentFlag, a.MeasurementPlanning
                        , d.CompanyId
                        , d.MtlItemNo, d.MtlItemName, d.MtlItemId, d.MtlItemSpec
                        , ISNULL(e.ProcessAlias, '') ProcessAlias
                        , ISNULL(f.StatusName, '') StatusName
                        , g.UserNo, g.UserName
                        , h.QcNoticeNo, h.QcNoticeQty
                        , i.StatusName CheckQcMeasureDataName
                        , (
                            SELECT (
                                SELECT xa.ControlId, xa.Cad2DFile, xa.Pdf2DFile
                                , xb.CustomerMtlItemNo, xb.CustomerCadControlId
                                , xc.CadFile
                                , xd.[FileName] + xd.FileExtension Cad2DFileFullName
                                , xe.[FileName] + xe.FileExtension CustomerCadFileFullName
                                , xf.[FileName] + xf.FileExtension Pdf2DFileFullName
                                FROM MES.RoutingItem x
                                INNER JOIN PDM.RdDesignControl xa ON x.ControlId = xa.ControlId
                                INNER JOIN PDM.RdDesign xb ON xa.DesignId = xb.DesignId
                                INNER JOIN PDM.CustomerCadControl xc ON xb.CustomerCadControlId = xc.ControlId
                                INNER JOIN BAS.[File] xd ON xa.Cad2DFile = xd.FileId
                                INNER JOIN BAS.[File] xe ON xc.CadFile = xe.FileId
                                LEFT JOIN BAS.[File] xf ON xa.Pdf2DFile = xf.FileId
                                WHERE x.RoutingItemId = j.RoutingItemId
                                FOR JSON PATH, ROOT('data')
                            ) CadInfoData
                            FROM MES.MoRouting j
                            WHERE j.MoId = a.MoId
                            FOR JSON PATH, ROOT('data')
                        ) CadInfo
                        , (
                            SELECT (
                                SELECT xa.Cad2DFileAbsolutePath, xa.Pdf2DFileAbsolutePath
                                , xb.CustomerMtlItemNo, xb.CustomerCadControlId
                                , xc.CadFileAbsolutePath
                                FROM MES.RoutingItem x
                                INNER JOIN PDM.RdDesignControl xa ON x.ControlId = xa.ControlId
                                INNER JOIN PDM.RdDesign xb ON xa.DesignId = xb.DesignId
                                INNER JOIN PDM.CustomerCadControl xc ON xb.CustomerCadControlId = xc.ControlId
                                WHERE x.RoutingItemId = j.RoutingItemId
                                FOR JSON PATH, ROOT('data')
                            ) CadInfoData
                            FROM MES.MoRouting j
                            WHERE j.MoId = a.MoId
                            FOR JSON PATH, ROOT('data')
                        ) CadFileAbsolutePathInfo
                        , k.QcTypeNo, k.QcTypeName, k.SupportAqFlag, k.SupportProcessFlag
                        , (
                              SELECT x.FileId
                              , xa.[FileName], xa.FileExtension
                              FROM MES.QcRecordFile x
                              INNER JOIN BAS.[File] xa ON x.FileId = xa.FileId
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'other'
                              AND x.FileId IS NOT NULL
                              FOR JSON PATH, ROOT('data')
                        ) QcRecordFile
                        , (
                              SELECT x.QcRecordFileId, x.PhysicalPath
                              FROM MES.QcRecordFile x
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'other'
                              AND x.PhysicalPath IS NOT NULL
                              FOR JSON PATH, ROOT('data')
                        ) QcRecordFileByNas
                        , (
                              SELECT x.FileId
                              , xa.[FileName] + xa.FileExtension FileFullName
                              FROM MES.QcRecordFile x
                              INNER JOIN BAS.[File] xa ON x.FileId = xa.FileId
                              WHERE x.QcRecordId = a.QcRecordId
                              AND x.FileType = 'resolve'
                              FOR JSON PATH, ROOT('data')
                        ) ResolveFile";
                    sqlQuery.mainTables =
                        @"FROM MES.QcRecord a
                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                        LEFT JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                        LEFT JOIN BAS.[Status] f ON a.QcStatus = f.StatusNo AND f.StatusSchema = 'QcRecord.QcStatus'
                        INNER JOIN BAS.[User] g ON a.CreateBy = g.UserId
                        LEFT JOIN QMS.QcNotice h ON a.QcNoticeId = h.QcNoticeId
                        INNER JOIN QMS.QcType k ON a.QcTypeId = k.QcTypeId
                        INNER JOIN BAS.Status i ON a.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND d.CompanyId = @CompanyId AND k.QcTypeNo = 'OIQC'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcRecordId", @" AND a.QcRecordId = @QcRecordId", QcRecordId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcNoticeNo", @" AND h.QcNoticeNo LIKE '%' + @QcNoticeNo + '%'", QcNoticeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND d.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND d.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND d.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcTypeId", @" AND a.QcTypeId = @QcTypeId", QcTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CheckQcMeasureData", @" AND a.CheckQcMeasureData = @CheckQcMeasureData", CheckQcMeasureData);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.CreateBy = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcType", @" AND k.QcTypeNo = @QcType", QcType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "CASE WHEN a.UrgentFlag = 'Y' THEN 0 WHEN a.CheckQcMeasureData = 'S' THEN 1 WHEN a.CheckQcMeasureData = 'C' THEN 2 WHEN a.CheckQcMeasureData = 'A' THEN 3 WHEN a.CheckQcMeasureData = 'N' THEN 4 WHEN a.CheckQcMeasureData = 'P' THEN 5 WHEN a.CheckQcMeasureData = 'Y' THEN 6 ELSE 7 END, a.QcRecordId DESC";
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

        #region //GetQcOutsourcingMeasureDataTemp -- 取得量測暫存資料  -- GPAI 2025-04-10
        public string GetQcOutsourcingMeasureDataTemp(int QcRecordId)
        {
            try
            {
                if (QcRecordId <= 0) throw new SystemException("量測單號不能為空!!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TempId, a.QcRecordId, a.Cell, a.Css, a.[Format], a.[Value]
                            FROM QMS.QcMeasureDataTemp a 
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    sql += @" ORDER BY a.TempId";

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

        #region //GetQcRecodDataQrcode -- 掃碼取得量測紀錄  -- GPAI 2025-04-29
        public string GetQcRecodDataQrcode(int QcRecordId, string baseUrl)
        {
            try
            {
                if (QcRecordId <= 0) throw new SystemException("量測單號不能為空!!");
                string QcPath = "";

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"select a.QcRecordId, b.QcTypeNo --金貨IQC 托外 OIQC
                            from MES.QcRecord a
                            INNER JOIN QMS.QcType b ON a.QcTypeId = b.QcTypeId 
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item in result)
                    {
                        switch (item.QcTypeNo)
                        {
                            case "IQC":
                                QcPath = $"{baseUrl}/QcDataManagement/QcGoodsReceiptManagement?QcRecordId={QcRecordId}";
                                break;
                            case "OIQC":
                                QcPath = $"{baseUrl}/QcDataManagement/QcOutsourcingManagement?QcRecordId={QcRecordId}";
                                break;
                            default:
                                QcPath = $"{baseUrl}/QcDataManagement/QcDataManagement?QcRecordId={QcRecordId}";
                                break;
                        }
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = QcPath
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
        #region //AddQcRecord -- 新增量測紀錄單頭 -- Ann 2023-02-27
        public string AddQcRecord(int QcNoticeId, int QcTypeId, int MoId, int MoProcessId, int QmmDetailId, int QcItemId, string Remark, int CurrentFileId, string SupportAqFlag, string LotNumber
            , string QcRecordFile, string InputType, string ServerPath, string ServerPath2, string ResolveFile, string ResolveFileJson, string QcRecordFileByNas, string QcMeasureDataJson)
        {
            try
            {
                string ErpDbName = "";
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (InputType.Length < 0) throw new SystemException("【輸入方式】不能為空!");
                            //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Cavity") throw new SystemException("【穴號模式】只能由全吋檢或成型工程檢使用!!");
                            //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Picture") throw new SystemException("【面型圖模式】只能由全吋檢或成型工程檢使用!!");

                            #region //判斷MES製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT (b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(varchar(10), a.WoSeq) + ')') WoErpFullNo
                                    , b.WoErpPrefix, b.WoErpNo
                                    , c.MtlItemId, c.MtlItemNo, c.MtlItemName
                                    FROM MES.ManufactureOrder a 
                                    INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    WHERE MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("MES製令資料錯誤!");

                            string WoErpFullNo = "";
                            int MtlItemId = -1;
                            string MtlItemNo = "";
                            string MtlItemName = "";
                            string WoErpPrefix = "";
                            string WoErpNo = "";
                            foreach (var item in result)
                            {
                                WoErpFullNo = item.WoErpFullNo;
                                MtlItemId = item.MtlItemId;
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                                WoErpPrefix = item.WoErpPrefix;
                                WoErpNo = item.WoErpNo;
                            }
                            #endregion

                            #region //判斷ERP製令狀態是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TA011, TA013
                                    FROM MOCTA
                                    WHERE TA001 = @TA001
                                    AND TA002 = @TA002";
                            dynamicParameters.Add("TA001", WoErpPrefix);
                            dynamicParameters.Add("TA002", WoErpNo);

                            var MOCTAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (MOCTAResult.Count() <= 0) throw new SystemException("ERP製令資料錯誤!!");

                            foreach (var item in MOCTAResult)
                            {
                                if (SupportAqFlag == "Y" && (item.TA011 == "Y" || item.TA011 == "y" || item.TA013 == "V")) throw new SystemException("ERP製令狀態無法開立!!");
                            }
                            #endregion

                            #region //判斷量測類型資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                    FROM QMS.QcType a 
                                    WHERE a.QcTypeId = @QcTypeId";
                            dynamicParameters.Add("QcTypeId", QcTypeId);

                            var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                            string QcTypeNo = "";
                            string QcTypeName = "";
                            string SupportProcessFlag = "";
                            foreach (var item in QcTypeResult)
                            {
                                if (item.SupportProcessFlag == "Y" && MoProcessId <= 0) throw new SystemException("【量測製程】不能為空!");
                                QcTypeNo = item.QcTypeNo;
                                QcTypeName = item.QcTypeName;
                                SupportProcessFlag = item.SupportProcessFlag;
                            }
                            #endregion

                            if (MoProcessId > 0)
                            {
                                #region //判斷製程資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.MoProcess
                                        WHERE MoProcessId = @MoProcessId";
                                dynamicParameters.Add("MoProcessId", MoProcessId);

                                var MoProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                if (MoProcessResult.Count() <= 0) throw new SystemException("製令製程資料錯誤!");
                                #endregion
                            }

                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            string DefaultSpreadsheetData = "";
                            foreach (var item in DefaultFileIdResult)
                            {
                                DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                            }
                            #endregion

                            #region //建立CheckQcMeasureData
                            string CheckQcMeasureData = "N";
                            if (InputType == "Cavity" || InputType == "Picture")
                            {
                                if (QcRecordFile.Length > 0)
                                {
                                    CheckQcMeasureData = "Y";
                                }
                            }
                            #endregion

                            #region //INSERT MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcRecord (QcNoticeId, QcTypeId, InputType, MoId, MoProcessId, Remark, DefaultFileId, CurrentFileId, DefaultSpreadsheetData, CheckQcMeasureData, SupportAqFlag, SupportProcessFlag, ResolveFileJson
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcRecordId
                                    VALUES (@QcNoticeId, @QcTypeId, @InputType, @MoId, @MoProcessId, @Remark, @DefaultFileId, @CurrentFileId, @DefaultSpreadsheetData, @CheckQcMeasureData, @SupportAqFlag, @SupportProcessFlag, @ResolveFileJson
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId = QcNoticeId <= 0 ? (int?)null : QcNoticeId,
                                    QcTypeId,
                                    InputType,
                                    MoId,
                                    MoProcessId = MoProcessId > 0 ? MoProcessId : (int?)null,
                                    Remark,
                                    DefaultFileId = -1,
                                    CurrentFileId,
                                    DefaultSpreadsheetData,
                                    CheckQcMeasureData,
                                    SupportAqFlag,
                                    SupportProcessFlag,
                                    ResolveFileJson,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int QcRecordId = -1;
                            foreach (var item in insertResult)
                            {
                                QcRecordId = item.QcRecordId;
                            }
                            #endregion

                            #region //INSERT MES.QcRecordFile
                            if (QcRecordFile.Length > 0)
                            {
                                var qcRecordFileList = QcRecordFile.Split(',');

                                foreach (var qcRecordFileId in qcRecordFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileId = qcRecordFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT MES.QcRecordFile By NAS
                            if (QcRecordFileByNas.Length > 0)
                            {
                                JObject uploadFileJson = JObject.Parse(QcRecordFileByNas);

                                foreach (var item in uploadFileJson["uploadFileInfo"])
                                {
                                    string fileType = item["FileType"]?.ToString() ?? "other";
                                    string inputType = item["InputType"]?.ToString() ?? null;
                                    string barcodeId = item["BarcodeId"]?.ToString() ?? null;

                                    //SplitCoating (SC) 分光鍍膜 file-management
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, InputType, BarcodeId, PhysicalPath, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @InputType, @BarcodeId, @PhysicalPath, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = fileType,
                                            InputType = inputType,
                                            BarcodeId = barcodeId,
                                            PhysicalPath = item["FilePath"].ToString(),
                                            LotNumber = LotNumber ?? null,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT QMS.QcMeasureData
                            if (QcMeasureDataJson.Length > 0)
                            {
                                if (string.IsNullOrWhiteSpace(LotNumber)) throw new SystemException("【批號】不能為空!");

                                #region //取得量測項目
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QcItemId, a.QcItemNo, a.QcItemName, a.QcType, ISNULL(b.QcItemDesc, a.QcItemDesc) QcItemDesc, b.*
                                        FROM QMS.QcItem a
                                        OUTER APPLY (
	                                        SELECT aa.QcItemDesc, aa.DesignValue, aa.UpperTolerance, aa.LowerTolerance, aa.SortNumber
	                                        , ab.QmmDetailId, ab.MachineNumber
	                                        , ac.MachineNo, ac.MachineName, ad.ShopName 
	                                        FROM PDM.MtlQcItem aa
	                                        INNER JOIN QMS.QmmDetail ab ON ab.QmmDetailId = aa.QmmDetailId
	                                        INNER JOIN MES.Machine ac ON ac.MachineId = ab.MachineId
	                                        INNER JOIN MES.WorkShop ad ON ad.ShopId = ac.ShopId AND ad.CompanyId = @CompanyId
	                                        WHERE aa.QcItemId = a.QcItemId
	                                        AND ad.CompanyId = @CompanyId
	                                        AND aa.MtlItemId = @MtlItemId
                                        ) b
                                        WHERE a.QcItemId = @QcItemId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("QcItemId", QcItemId);
                                var resultQcItem = sqlConnection.Query(sql, dynamicParameters);
                                if (!resultQcItem.Any()) throw new SystemException("【量測項目】資料錯誤!");
                                string QcItemDesc = string.Empty;
                                string DesignValue = string.Empty;
                                string UpperTolerance = string.Empty;
                                string LowerTolerance = string.Empty;
                                foreach (var item in resultQcItem)
                                {
                                    QcItemDesc = item.QcItemDesc;
                                    DesignValue = item.DesignValue ?? "OK";
                                    UpperTolerance = item.UpperTolerance;
                                    LowerTolerance = item.LowerTolerance;
                                    QmmDetailId = item.QmmDetailId ?? QmmDetailId;
                                }
                                #endregion

                                #region //取得批號 (停用)
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT a.LotNumberId, a.MtlItemId, a.LotNumberNo
                                //        FROM SCM.LotNumber a
                                //        WHERE a.MtlItemId = @MtlItemId";
                                //dynamicParameters.Add("MtlItemId", MtlItemId);
                                //var resultLotNumber = sqlConnection.Query(sql, dynamicParameters);
                                //string LotNumber = string.Empty;
                                //foreach (var item in resultLotNumber)
                                //{
                                //    LotNumber = item.LotNumberNo;
                                //}
                                #endregion

                                JObject qcMeasureData = JObject.Parse(QcMeasureDataJson);

                                foreach (var item in qcMeasureData["data"])
                                {
                                    int.TryParse(item["BarcodeId"]?.ToString(), out int BarcodeId);
                                    string MeasureValue = item["MeasureValue"]?.ToString() ?? "";

                                    #region //新增量測資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO QMS.QcMeasureData (QcRecordId, QcItemId, QcItemDesc, DesignValue, BarcodeId, QmmDetailId, MeasureValue, QcResult, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QmdId
                                            VALUES (@QcRecordId, @QcItemId, @QcItemDesc, @DesignValue, @BarcodeId, @QmmDetailId, @MeasureValue, @QcResult, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            QcItemId,
                                            QcItemDesc,
                                            DesignValue,
                                            BarcodeId,
                                            QmmDetailId,
                                            MeasureValue,
                                            QcResult = MeasureValue == "OK" ? "P" : "F",
                                            LotNumber = LotNumber ?? null,
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
                            #endregion

                            #region //INSERT ResolveFile
                            if (ResolveFile.Length > 0)
                            {
                                var resolveFileList = ResolveFile.Split(',');

                                foreach (var resolveFileId in resolveFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = "resolve",
                                            FileId = resolveFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //建立量測單據Spreadsheet DATA
                            #region //取得品號允收標準
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                    , a.QcItemDesc, a.SortNumber, a.QmmDetailId, a.BallMark, a.Unit, a.QmmDetailId
                                    , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                    , c.MachineNumber
                                    , d.MachineDesc
                                    FROM PDM.MtlQcItem a
                                    INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                    LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                    LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.[Status] = 'A' 
                                    ORDER BY a.SortNumber";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //初始化Data
                            List<Data> datas = new List<Data>();
                            Data data = new Data();
                            data = new Data()
                            {
                                cell = "A1",
                                css = "imported_class1",
                                format = "text",
                                value = "序號",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "B1",
                                css = "imported_class2",
                                format = "text",
                                value = "球標",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "C1",
                                css = "imported_class2",
                                format = "text",
                                value = "檢測項目",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "D1",
                                css = "imported_class3",
                                format = "text",
                                value = "檢測備註",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "E1",
                                css = "imported_class4",
                                format = "text",
                                value = "量測設備",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "F1",
                                css = "imported_class5",
                                format = "text",
                                value = "設計值",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "G1",
                                css = "imported_class6",
                                format = "text",
                                value = "上限",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "H1",
                                css = "imported_class7",
                                format = "text",
                                value = "下限",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "I1",
                                css = "imported_class8",
                                format = "text",
                                value = "單位",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "J1",
                                css = "imported_class8",
                                format = "text",
                                value = "Z軸",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "K1",
                                css = "imported_class9",
                                format = "text",
                                value = "量測人員",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "L1",
                                css = "imported_class10",
                                format = "text",
                                value = "量測工時",
                            };
                            datas.Add(data);
                            #endregion

                            #region //設定單身量測標準
                            int row = 2;
                            foreach (var item2 in MtlQcItemResult)
                            {
                                #region //若有機台，整理序號格式
                                string QcItemNo = item2.QcItemNo;
                                if (item2.MachineNumber != null)
                                {
                                    QcItemNo = item2.QcItemNo;
                                    string firstPart = QcItemNo.Substring(0, 3);
                                    string secondPart = QcItemNo.Substring(3, 4);
                                    QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                }
                                #endregion

                                string QcItemName = item2.QcItemName;

                                #region //設定量測項目、備註、上下限
                                data = new Data()
                                {
                                    cell = "A" + row,
                                    css = "",
                                    format = "common",
                                    value = QcItemNo,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.BallMark != null ? item2.BallMark.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.QcItemName,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.QcItemDesc,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "H" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "I" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.Unit != null ? item2.Unit.ToString() : "",
                                };
                                datas.Add(data);

                                row++;
                                #endregion
                            }
                            #endregion

                            #region //整合Spreadsheet格式
                            List<Sheets> sheetss = new List<Sheets>();

                            Sheets sheets = new Sheets()
                            {
                                name = "sheet1",
                                data = datas
                            };
                            sheetss.Add(sheets);

                            #region //更新至QcRecord SpreadsheetData
                            SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                            {
                                sheets = sheetss
                            };

                            string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    SpreadsheetData = @SpreadsheetData,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SpreadsheetData,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion
                            #endregion

                            #region //將量測EXCEL先存回實體路徑並重新讀取
                            string FileName = "";
                            if (CurrentFileId > 0)
                            {
                                #region //取得原本加密檔案資料，並重新上傳
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId, a.FileContent, a.FileName, a.FileExtension, a.FileSize, a.ClientIP, a.Source
                                        FROM BAS.[File] a
                                        WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", CurrentFileId);

                                var FileResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in FileResult)
                                {
                                    FileName = item.FileName;
                                    ServerPath = Path.Combine(ServerPath, item.FileName + "-" + item.FileId + item.FileExtension);
                                    byte[] fileContent = (byte[])item.FileContent;
                                    File.WriteAllBytes(ServerPath, fileContent); // Requires System.IO

                                    #region //將檔案重新上傳
                                    System.Net.WebClient wc = new System.Net.WebClient();
                                    byte[] newFilebytes = wc.DownloadData(ServerPath);

                                    dynamicParameters = new DynamicParameters();
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
                                            CompanyId = CurrentCompany,
                                            item.FileName,
                                            FileContent = newFilebytes,
                                            item.FileExtension,
                                            item.FileSize,
                                            item.ClientIP,
                                            item.Source,
                                            DeleteStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    int newFileId = -1;
                                    foreach (var item2 in insertResult)
                                    {
                                        newFileId = item2.FileId;
                                    }
                                    #endregion

                                    #region //UPDATE原本FileId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcRecord SET
                                            CurrentFileId = @CurrentFileId,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CurrentFileId = newFileId,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QcRecordId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除實體檔案
                                    File.Delete(ServerPath);
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //統整回傳前端資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId, a.CurrentFileId
                                    , (
                                        SELECT x.FileId
                                        FROM MES.QcRecordFile x
                                        WHERE x.QcRecordId = a.QcRecordId
                                        FOR JSON PATH, ROOT('data')
                                    ) QcRecordFile
                                    FROM MES.QcRecord a
                                    WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            if ((CheckQcMeasureData == "Y" || CheckQcMeasureData == "P") && CurrentFileId > 0 && (QcTypeNo == "PVTQC" || QcTypeNo == "TQC"))
                            {
                                SendQcMail(sqlConnection, QcRecordId, WoErpFullNo, MtlItemNo, MtlItemName, QcTypeNo, QcTypeName
                                    , Remark, FileName, ServerPath2, MoId);
                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                data = QcRecordResult
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

        #region //AddQcRecordFile -- 新增量測檔案(量測檔案歸檔功能) -- Ann 2024-01-04
        public string AddQcRecordFile(int QcRecordId, string InputType, string FileInfo)
        {
            try
            {
                if (FileInfo.Length <= 0) throw new SystemException("量測檔案資料錯誤!!");
                if (InputType.Length <= 0) throw new SystemException("歸檔方式錯誤!!");

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測單據資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoId
                                , c.MtlItemId
                                FROM MES.QcRecord a 
                                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測單據資料錯誤!!");

                        int MoId = -1;
                        int MtlItemId = -1;
                        foreach (var item in QcRecordResult)
                        {
                            MoId = item.MoId;
                            MtlItemId = item.MtlItemId;
                        }
                        #endregion

                        #region //先刪除原先全部歸檔檔案
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcRecordFile
                                WHERE QcRecordId = @QcRecordId
                                AND FileType = 'file-management'";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        var fileInfoJson = JObject.Parse(FileInfo);

                        int? BarcodeId = null;
                        string LotNumber = null;
                        string Cavity = null;
                        string EffectiveDiameter = null;
                        foreach (var item in fileInfoJson["datas"])
                        {
                            switch (InputType)
                            {
                                #region //條碼歸檔
                                case "BarcodeNo":
                                    #region //確認條碼資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT BarcodeId
                                            FROM MES.Barcode a 
                                            WHERE a.BarcodeNo = @BarcodeNo
                                            AND a.MoId = @MoId";
                                    dynamicParameters.Add("BarcodeNo", item["InputValue"].ToString());
                                    dynamicParameters.Add("MoId", MoId);

                                    var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (BarcodeResult.Count() <= 0) throw new SystemException("條碼【" + item["InputValue"].ToString() + "】資料錯誤!!");

                                    foreach (var item2 in BarcodeResult)
                                    {
                                        BarcodeId = item2.BarcodeId;
                                    }
                                    #endregion
                                    break;
                                #endregion
                                #region //刻字歸檔
                                case "Lettering":
                                    #region //確認條碼資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.BarcodeId
                                            FROM MES.BarcodeAttribute a 
                                            INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                            WHERE a.ItemNo = 'Lettering'
                                            ANd a.ItemValue = @ItemValue
                                            AND b.MoId = @MoId";
                                    dynamicParameters.Add("MoId", MoId);
                                    dynamicParameters.Add("ItemValue", item["InputValue"].ToString());

                                    var BarcodeAttributeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (BarcodeAttributeResult.Count() <= 0) throw new SystemException("刻字流水號【" + item["InputValue"].ToString() + "】資料錯誤!!");

                                    foreach (var item2 in BarcodeAttributeResult)
                                    {
                                        BarcodeId = item2.BarcodeId;
                                    }
                                    #endregion
                                    break;
                                #endregion
                                #region //批號歸檔
                                case "LotNumber":
                                    #region //確認條碼資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.LotNumberNo
                                            FROM SCM.LotNumber a 
                                            WHERE a.MtlItemId = @MtlItemId 
                                            AND a.LotNumberNo = @LotNumber";
                                    dynamicParameters.Add("MtlItemId", MtlItemId);
                                    dynamicParameters.Add("LotNumber", item["InputValue"].ToString());

                                    var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (LotNumberResult.Count() <= 0) throw new SystemException("批號【" + item["InputValue"].ToString() + "】資料錯誤!!");

                                    foreach (var item2 in LotNumberResult)
                                    {
                                        LotNumber = item2.LotNumberNo;
                                    }
                                    #endregion

                                    #region //確認R1 R2資料
                                    if (item["EffectiveDiameter"].ToString().Length <= 0) throw new SystemException("R1R2資料錯誤!!");
                                    EffectiveDiameter = item["EffectiveDiameter"].ToString();
                                    #endregion
                                    break;
                                #endregion
                                #region //穴號歸檔
                                case "Cavity":
                                    #region //格式檢查
                                    Cavity = item["InputValue"].ToString();
                                    if (Cavity.IndexOf("-") == -1) throw new SystemException("穴號【" + item["InputValue"].ToString() + "】格式錯誤!!");
                                    #endregion

                                    #region //確認R1 R2資料
                                    if (item["EffectiveDiameter"].ToString().Length <= 0) throw new SystemException("R1R2資料錯誤!!");
                                    EffectiveDiameter = item["EffectiveDiameter"].ToString();
                                    #endregion
                                    break;
                                #endregion
                                #region /Catch
                                default:
                                    throw new SystemException("歸檔方式錯誤!!");
                                    #endregion
                            }

                            #region //INSERT MES.QcRecordFile
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, PhysicalPath, InputType, BarcodeId, Cavity, LotNumber, EffectiveDiameter
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcRecordFileId
                                    VALUES (@QcRecordId, @FileType, @PhysicalPath, @InputType, @BarcodeId, @Cavity, @LotNumber, @EffectiveDiameter
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcRecordId,
                                    FileType = "file-management",
                                    PhysicalPath = item["FilePath"].ToString(),
                                    InputType,
                                    BarcodeId,
                                    Cavity,
                                    LotNumber,
                                    EffectiveDiameter,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = QcRecordResult
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

        #region //AddQcGoodsReceipt -- 新增進貨檢量測單據 -- Ann 2024-04-02
        public string AddQcGoodsReceipt(int GrDetailId, string InputType, string QcRecordFile, string ResolveFile, string ResolveFileJson, string Remark, string ServerPath, string ServerPath2)
        {
            try
            {
                string ErpDbName = "";
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (InputType.Length < 0) throw new SystemException("【輸入方式】不能為空!");

                            #region //確認進貨單身資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.CloseStatus, a.MtlItemId
                                    , b.ConfirmStatus
                                    FROM SCM.GrDetail a 
                                    INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");

                            int MtlItemId = -1;
                            foreach (var item in GrDetailResult)
                            {
                                if (item.CloseStatus != "N") throw new SystemException("此進貨單身已結案!!");
                                if (item.ConfirmStatus != "N") throw new SystemException("此進貨單身已確認!!");

                                MtlItemId = item.MtlItemId;
                            }
                            #endregion

                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            string DefaultSpreadsheetData = "";
                            foreach (var item in DefaultFileIdResult)
                            {
                                DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                            }
                            #endregion

                            #region //取得進貨檢QcType
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 QcTypeId
                                    FROM QMS.QcType
                                    WHERE QcTypeNo = 'IQC'";

                            var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            int QcTypeId = -1;
                            foreach (var item in QcTypeResult)
                            {
                                QcTypeId = item.QcTypeId;
                            }
                            #endregion

                            #region //確認此進貨單身是否有其他量測單據量測
                            dynamicParameters = new DynamicParameters();
                            sql = @"";
                            #endregion

                            #region //確認此進貨單身是否有品異單尚未判定完成
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.AbnormalqualityNo
                                    FROM QMS.AqBarcode a 
                                    INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                    WHERE b.GrDetailId = @GrDetailId
                                    AND a.JudgeStatus IS NULL";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in AqBarcodeResult)
                            {
                                throw new SystemException("此進貨單身尚有品異單【" + item.AbnormalqualityNo + "】未判定完成!!");
                            }
                            #endregion

                            #region //INSERT MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, Remark, DefaultFileId, DefaultSpreadsheetData, ResolveFileJson
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcRecordId
                                    VALUES (@QcTypeId, @InputType, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @ResolveFileJson
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcTypeId,
                                    InputType,
                                    Remark,
                                    DefaultFileId = -1,
                                    DefaultSpreadsheetData,
                                    ResolveFileJson,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int QcRecordId = -1;
                            foreach (var item in insertResult)
                            {
                                QcRecordId = item.QcRecordId;
                            }
                            #endregion

                            #region //INSERT MES.QcGoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcGoodsReceipt (QcRecordId, GrDetailId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcGoodsReceiptId
                                    VALUES (@QcRecordId, @GrDetailId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcRecordId,
                                    GrDetailId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INSERT MES.QcRecordFile
                            if (QcRecordFile.Length > 0)
                            {
                                var qcRecordFileList = QcRecordFile.Split(',');

                                foreach (var qcRecordFileId in qcRecordFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileId = qcRecordFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT ResolveFile
                            if (ResolveFile.Length > 0)
                            {
                                var resolveFileList = ResolveFile.Split(',');

                                foreach (var resolveFileId in resolveFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = "resolve",
                                            FileId = resolveFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //建立量測單據Spreadsheet DATA
                            #region //取得品號允收標準
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                    , a.QcItemDesc, a.SortNumber, a.QmmDetailId
                                    , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                    , c.MachineNumber
                                    , d.MachineDesc
                                    FROM PDM.MtlQcItem a
                                    INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                    LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                    LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.[Status] = 'A' 
                                    ORDER BY a.SortNumber";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //初始化Data
                            List<Data> datas = new List<Data>();
                            Data data = new Data();
                            data = new Data()
                            {
                                cell = "A1",
                                css = "imported_class1",
                                format = "common",
                                value = "序號",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "B1",
                                css = "imported_class2",
                                format = "common",
                                value = "檢測項目",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "C1",
                                css = "imported_class3",
                                format = "common",
                                value = "檢測備註",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "D1",
                                css = "imported_class4",
                                format = "common",
                                value = "量測設備",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "E1",
                                css = "imported_class5",
                                format = "common",
                                value = "設計值",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "F1",
                                css = "imported_class6",
                                format = "common",
                                value = "上限",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "G1",
                                css = "imported_class7",
                                format = "common",
                                value = "下限",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "H1",
                                css = "imported_class8",
                                format = "common",
                                value = "Z軸",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "I1",
                                css = "imported_class9",
                                format = "common",
                                value = "量測人員",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "J1",
                                css = "imported_class10",
                                format = "common",
                                value = "量測工時",
                            };
                            datas.Add(data);
                            #endregion

                            #region //設定單身量測標準
                            int row = 2;
                            foreach (var item2 in MtlQcItemResult)
                            {
                                #region //若有機台，整理序號格式
                                string QcItemNo = item2.QcItemNo;
                                if (item2.MachineNumber != null)
                                {
                                    QcItemNo = item2.QcItemNo;
                                    string firstPart = QcItemNo.Substring(0, 3);
                                    string secondPart = QcItemNo.Substring(3, 4);
                                    QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                }
                                #endregion

                                string QcItemName = item2.QcItemName;

                                #region //設定量測項目、備註、上下限
                                data = new Data()
                                {
                                    cell = "A" + row,
                                    css = "",
                                    format = "common",
                                    value = QcItemNo,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B" + row,
                                    css = "",
                                    format = "common",
                                    value = QcItemName,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C" + row,
                                    css = "",
                                    format = "common",
                                    value = item2.QcItemDesc,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D" + row,
                                    css = "",
                                    format = "common",
                                    value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E" + row,
                                    css = "",
                                    format = "common",
                                    value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F" + row,
                                    css = "",
                                    format = "common",
                                    value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G" + row,
                                    css = "",
                                    format = "common",
                                    value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                };
                                datas.Add(data);

                                row++;
                                #endregion
                            }
                            #endregion

                            #region //整合Spreadsheet格式
                            List<Sheets> sheetss = new List<Sheets>();

                            Sheets sheets = new Sheets()
                            {
                                name = "sheet1",
                                data = datas
                            };
                            sheetss.Add(sheets);

                            #region //更新至QcRecord SpreadsheetData
                            SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                            {
                                sheets = sheetss
                            };

                            string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    SpreadsheetData = @SpreadsheetData,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SpreadsheetData,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion
                            #endregion

                            #region //統整回傳前端資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId, a.CurrentFileId
                                    , (
                                        SELECT x.FileId
                                        FROM MES.QcRecordFile x
                                        WHERE x.QcRecordId = a.QcRecordId
                                        FOR JSON PATH, ROOT('data')
                                    ) QcRecordFile
                                    FROM MES.QcRecord a
                                    WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                data = QcRecordResult
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

        #region //AddQcRecordPlanning -- 新增量測單據排程資料 -- Ann 2024-08-20
        public string AddQcRecordPlanning(int QcRecordId, string UploadJson, string Spreadsheet)
        {
            try
            {
                if (UploadJson.Length < 0) throw new SystemException("【量測機台彙整資料】不能為空!!");

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
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!!");

                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion


                        #region //確認量測單據資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CheckQcMeasureData
                                FROM MES.QcRecord a 
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測單據資料錯誤!!");

                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "A" && item.CheckQcMeasureData != "C")
                            {
                                throw new SystemException("量測單據狀態錯誤!!");
                            }
                        }
                        #endregion

                        #region //解析量測機型數據
                        JObject uploadJson = JObject.Parse(UploadJson);
                        if (uploadJson["uploadInfo"][0].Count() <= 0) throw new SystemException("量測機型排程數據資料錯誤!!");
                        Dictionary<int, int> QcMachineModeDic = new Dictionary<int, int>();
                        Dictionary<int, string> QcItemNoDic = new Dictionary<int, string>();

                        foreach (var item in uploadJson["uploadInfo"])
                        {
                            string QcItemNo = item["QcItemNo"].ToString();
                            int QcMachineModeId = Convert.ToInt32(item["QcMachineMode"].ToString().Split(':')[0]);

                            #region //確認量測項目是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcItem a 
                                    INNER JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                                    INNER JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                                    WHERE a.QcItemNo = @QcItemNo
                                    AND c.CompanyId = @CompanyId";
                            dynamicParameters.Add("QcItemNo", QcItemNo);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var QcItemResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcItemResult.Count() <= 0) throw new SystemException("【量測項目: " + QcItemNo + "】資料錯誤!!");
                            if (!QcItemNoDic.ContainsKey(QcMachineModeId))
                            {
                                QcItemNoDic.Add(QcMachineModeId, QcItemNo);
                            }
                            #endregion

                            #region //確認量測機型是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcMachineMode a 
                                    WHERE a.QcMachineModeId = @QcMachineModeId";
                            dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                            var QcMachineModeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcMachineModeResult.Count() <= 0) throw new SystemException("【量測機型編號: " + QcMachineModeId + "】資料錯誤!!");
                            #endregion

                            if (!QcMachineModeDic.ContainsKey(QcMachineModeId))
                            {
                                QcMachineModeDic.Add(QcMachineModeId, 1);
                            }
                            else
                            {
                                int oriQcItemCount = QcMachineModeDic[QcMachineModeId];
                                QcMachineModeDic[QcMachineModeId] = oriQcItemCount + 1;
                            }
                        }
                        #endregion

                        #region //確認此量測單據是否已經有排程記錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.QcMachineModeName 
                                , c.QmmpId
                                , g.TypeThree, g.MtlItemNo, g.MtlItemId
                                FROM QMS.QcRecordPlanning a 
                                INNER JOIN QMS.QcMachineMode b ON a.QcMachineModeId = b.QcMachineModeId
                                LEFT JOIN QMS.QcMachineModePlanning c ON a.QrpId = c.QrpId
                                INNER JOIN MES.QcRecord d ON a.QcRecordId = d.QcRecordId
							    INNER JOIN MES.ManufactureOrder e ON d.MoId = e.MoId
                                INNER JOIN MES.WipOrder f ON e.WoId = f.WoId
                                INNER JOIN PDM.MtlItem g ON f.MtlItemId = g.MtlItemId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordPlanningResult = sqlConnection.Query(sql, dynamicParameters);
                        string mtlType = "";
                        string mtlTypeName = "";
                        

                        foreach (var item2 in QcRecordPlanningResult)
                        {
                            if (item2.QmmpId != null)
                            {
                                throw new SystemException("量測機型【" + item2.QcMachineModeName + "】已有排定排程!!");
                            }
                            mtlType = item2.TypeThree;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"SELECT LTRIM(RTRIM(MA002)) TypeNo, LTRIM(RTRIM(MA003)) TypeName
                                    FROM INVMA 
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeNo", @" AND LTRIM(RTRIM(MA002)) = @TypeNo", mtlType);
                            var mtlTyperesult = sqlConnection2.Query(sql, dynamicParameters);
                            if (mtlTyperesult.Count() <= 0) throw new SystemException("ERP品號類別資料錯誤!!");

                            foreach (var item in mtlTyperesult)
                            {

                                mtlTypeName = item.TypeName;
                            }

                        }

                        #region //刪除原本單據機型資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcRecordPlanning
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //INSERT QMS.QcRecordPlanning
                        foreach (var item in QcMachineModeDic)
                        {

                            #region //取得量測項目工時總和
                            string QicNo = QcItemNoDic[item.Key].Substring(3, 2);
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(CONVERT(DECIMAL(10,2), a.WorkTime/60.0)) AS TotalWorkTime
                                    FROM QMS.QcMeasureInTheoryWorkTime a
                                    INNER JOIN QMS.QmmDetail b ON a.QmmDetailId = b.QmmDetailId
                                    INNER JOIN MES.Machine c ON b.MachineId = c.MachineId
                                    INNER JOIN QMS.QcItemCoding d ON a.QicId = d.QicId
                                    WHERE b.QcMachineModeId = (
                                        SELECT QcMachineModeId 
                                        FROM QMS.QcMachineMode 
                                        WHERE QcMachineModeId = @QcMachineModeId
                                    )
                                    AND d.QicNo = @QicNo
                                    AND a.ProductType = @MtlTypeName
                                    AND a.QmwtId IN (
                                        SELECT MAX(QmwtId)
                                        FROM QMS.QcMeasureInTheoryWorkTime
                                        WHERE QicId IN (
                                            SELECT QicId 
                                            FROM QMS.QcItemCoding 
                                            WHERE QicNo = @QicNo
                                        )
                                        GROUP BY MeasureSize
                                    );";
                            dynamicParameters.Add("QcMachineModeId", item.Key);
                            dynamicParameters.Add("QicNo", QicNo);
                            dynamicParameters.Add("MtlTypeName", mtlTypeName);

                            var QcMeasureInTheoryWorkTimeResult = sqlConnection.Query(sql, dynamicParameters);
                            decimal estimatedMeasurementTime = 0.0M;

                            foreach (var item2 in QcMeasureInTheoryWorkTimeResult)
                            {
                                if (item2.TotalWorkTime == null || item2.TotalWorkTime <= 0)
                                {
                                    estimatedMeasurementTime = 0;
                                }
                                else
                                {
                                    estimatedMeasurementTime = item2.TotalWorkTime;
                                }
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.QcRecordPlanning (QcRecordId, QcMachineModeId, TotalQcItemCount, EstimatedMeasurementTime
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@QcRecordId, @QcMachineModeId, @TotalQcItemCount, @EstimatedMeasurementTime
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcRecordId,
                                    QcMachineModeId = item.Key,
                                    TotalQcItemCount = item.Value,
                                    EstimatedMeasurementTime = estimatedMeasurementTime,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //UPDATE MES.QcRecord PlanningSpreadsheetDate
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET 
                                PlanningSpreadsheetDate = @PlanningSpreadsheetDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PlanningSpreadsheetDate = Spreadsheet,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //AddQcMachineModePlanning -- 新增量測機型排程資料 -- Ann 2024-08-21
        public string AddQcMachineModePlanning(int QrpId, int QmmDetailId, int Sort)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認量測單據排程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId, a.EstimatedMeasurementTime, a.QcMachineModeId, a.ConfirmStatus
                                , b.CheckQcMeasureData
                                FROM QMS.QcRecordPlanning a 
                                INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                WHERE QrpId = @QrpId";
                        dynamicParameters.Add("QrpId", QrpId);

                        var QcRecordPlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcRecordPlanningResult.Count() <= 0) throw new SystemException("量測機型排程資料錯誤!!");

                        int QcRecordId = -1;
                        double EstimatedMeasurementTime = 0;
                        int QcMachineModeId = -1;
                        foreach (var item in QcRecordPlanningResult)
                        {
                            if (item.ConfirmStatus != "N")
                            {
                                throw new SystemException("量測單據排程已確認，無法新增!!");
                            }

                            if (item.EstimatedMeasurementTime <= 0)
                            {
                                throw new SystemException("預估工時尚未維護，無法進行排程!!");
                            }

                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "A" && item.CheckQcMeasureData != "C")
                            {
                                throw new SystemException("量測單據狀態錯誤!!");
                            }

                            QcRecordId = item.QcRecordId;
                            EstimatedMeasurementTime = Convert.ToDouble(item.EstimatedMeasurementTime);
                            QcMachineModeId = item.QcMachineModeId;
                        }
                        #endregion

                        #region //確認機台資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcMachineModeId
                                FROM QMS.QmmDetail a 
                                WHERE a.QmmDetailId = @QmmDetailId";
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

                        var QmmDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QmmDetailResult.Count() <= 0) throw new SystemException("量測機台資料錯誤!!");

                        foreach (var item in QmmDetailResult)
                        {
                            if (item.QcMachineModeId != QcMachineModeId) throw new SystemException("量測機台與量測機型對應錯誤!!");
                        }
                        #endregion

                        #region //確認此量測單據機型是否已排定排程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 FORMAT(a.EstimatedEndDate, 'yyyy-MM-dd HH:mm:ss') EstimatedEndDate 
                                , c.MachineDesc
                                , d.QcRecordId
                                FROM QMS.QcMachineModePlanning a 
                                INNER JOIN QMS.QmmDetail b ON a.QmmDetailId = b.QmmDetailId
                                INNER JOIN MES.Machine c ON b.MachineId = c.MachineId
                                INNER JOIN QMS.QcRecordPlanning d ON a.QrpId = d.QrpId
                                WHERE a.QrpId = @QrpId
                                AND a.QmmDetailId = @QmmDetailId";
                        dynamicParameters.Add("QrpId", QrpId);
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

                        var QrpIdCheckResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in QrpIdCheckResult)
                        {
                            throw new SystemException("量測單據【" + QcRecordId + "】<br>" +
                                "機台【" + item.MachineDesc + "】<br>" +
                                "已排定排程，預計【" + item.EstimatedEndDate + "】結束，無法重複排定!!");
                        }
                        #endregion

                        #region //確定順序是否為連續號碼
                        if (Sort != 1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcMachineModePlanning a 
                                    WHERE a.Sort = @Sort
                                    AND a.QmmDetailId = @QmmDetailId
                                    AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow";
                            dynamicParameters.Add("Sort", Sort - 1);
                            dynamicParameters.Add("QmmDetailId", QmmDetailId);
                            dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                            var PlanningSortCheckResult = sqlConnection.Query(sql, dynamicParameters);

                            if (PlanningSortCheckResult.Count() <= 0) throw new SystemException("【排序: " + Sort + "】非連續號碼，請重新確認後再輸入!!");
                        }
                        #endregion

                        #region //排程計算核心邏輯

                        #region //取得此次新排程初始開始日期
                        DateTime firstStartDate = DateTime.Now;
                        if (Sort != 1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.EstimatedEndDate
                                    FROM QMS.QcMachineModePlanning a 
                                    WHERE a.Sort = @Sort
                                    AND a.QmmDetailId = @QmmDetailId
                                    AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow";
                            dynamicParameters.Add("Sort", Sort - 1);
                            dynamicParameters.Add("QmmDetailId", QmmDetailId);
                            dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                            var EstimatedEndDateResult = sqlConnection.Query(sql, dynamicParameters);

                            if (EstimatedEndDateResult.Count() <= 0) throw new SystemException("取得初始開始日期時錯誤!!");

                            foreach (var item in EstimatedEndDateResult)
                            {
                                firstStartDate = item.EstimatedEndDate;
                            }
                        }
                        #endregion

                        #region //確認此量測單據是否已經有其他機型已排定排程，若有則需檢核時間順序
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.QmmDetailId, a.EstimatedEndDate
                                , d.MachineDesc
                                FROM QMS.QcMachineModePlanning a 
                                INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                INNER JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                INNER JOIN MES.Machine d ON c.MachineId = d.MachineId
                                WHERE b.QcRecordId = @QcRecordId
                                AND a.QrpId != @QrpId
                                AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow 
                                ORDER BY a.EstimatedEndDate DESC";
                        dynamicParameters.Add("QcRecordId", QcRecordId);
                        dynamicParameters.Add("QrpId", QrpId);
                        dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                        var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in QcMachineModePlanningResult)
                        {
                            DateTime finalEndDate = item.EstimatedEndDate;
                            if (firstStartDate < finalEndDate)
                            {
                                firstStartDate = finalEndDate;
                                //throw new SystemException("此次排程開始時間【" + firstStartDate + "】不可早於上個排程【" + item.MachineDesc + "】結束時間【" + finalEndDate.ToString("yyyy-MM-dd HH:mm:ss") + "】!!");
                            }
                        }
                        #endregion

                        #region //取得新排程之後的所有異動排程資料
                        List<int> OriQmmpIdSortArray = new List<int>();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QmmpId, a.Sort, a.EstimatedStartDate, a.EstimatedEndDate
                                , b.EstimatedMeasurementTime
                                , c.QcStartDate, c.QcRecordId
                                FROM QMS.QcMachineModePlanning a 
                                INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                INNER JOIN MES.QcRecord c ON b.QcRecordId = c.QcRecordId
                                WHERE a.Sort >= @Sort
                                AND a.QmmDetailId = @QmmDetailId
                                AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow
                                ORDER BY a.Sort";
                        dynamicParameters.Add("Sort", Sort);
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);
                        dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                        var PlanningSortCheckResult2 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in PlanningSortCheckResult2)
                        {
                            if (item.QcStartDate != null)
                            {
                                throw new SystemException("【原排序: " + item.Sort + "】已開始量測，無法變更其排程順序!!");
                            }

                            OriQmmpIdSortArray.Add(item.QmmpId);
                        }
                        #endregion

                        #region //Remove原先排程順序
                        if (OriQmmpIdSortArray.Count() > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                    Sort = NULL,
                                    EstimatedStartDate = NULL,
                                    EstimatedEndDate = NULL
                                    WHERE QmmpId IN (" + string.Join(",", OriQmmpIdSortArray) + ")";

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //先排序新排程
                        DateTime EstimatedEndDate = firstStartDate.AddMinutes(EstimatedMeasurementTime);
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcMachineModePlanning (QrpId, QmmDetailId, EstimatedStartDate, EstimatedEndDate, Sort
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@QrpId, @QmmDetailId, @EstimatedStartDate, @EstimatedEndDate, @Sort
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QrpId,
                                QmmDetailId,
                                EstimatedStartDate = firstStartDate,
                                EstimatedEndDate,
                                Sort,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Reset原先排程順序
                        int i = 1;
                        DateTime currentEstimatedEndDate = EstimatedEndDate;
                        foreach (var qmmpId in OriQmmpIdSortArray)
                        {
                            #region //計算新的開始日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.EstimatedMeasurementTime
                                    FROM QMS.QcMachineModePlanning a 
                                    INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                    WHERE a.QmmpId = @QmmpId";
                            dynamicParameters.Add("QmmpId", qmmpId);

                            var GetEstimatedMeasurementTimeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GetEstimatedMeasurementTimeResult.Count() <= 0) throw new SystemException("【量測機台排程ID: " + qmmpId + "】查無此排程資料!!");

                            double currentEstimatedMeasurementTime = 0;
                            foreach (var item in GetEstimatedMeasurementTimeResult)
                            {
                                currentEstimatedMeasurementTime = item.EstimatedMeasurementTime;
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                    Sort = @Sort,
                                    EstimatedStartDate = @EstimatedStartDate,
                                    EstimatedEndDate = @EstimatedEndDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QmmpId = @QmmpId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Sort = Sort + i,
                                    EstimatedStartDate = currentEstimatedEndDate,
                                    EstimatedEndDate = currentEstimatedEndDate.AddMinutes(currentEstimatedMeasurementTime),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QmmpId = qmmpId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            currentEstimatedEndDate = EstimatedEndDate.AddMinutes(currentEstimatedMeasurementTime);

                            i++;
                        }
                        #endregion

                        #endregion

                        #region //確認此量測單所有機型是否都已排程完畢
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcRecordPlanning a 
                                WHERE a.QcRecordId = @QcRecordId 
                                AND NOT EXISTS (
                                    SELECT TOP 1 1 FROM QMS.QcMachineModePlanning x 
                                    WHERE x.QrpId = a.QrpId
                                )";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordPlanningCheckResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcRecordPlanningCheckResult.Count() <= 0)
                        {
                            #region //更改量測單據排程狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET 
                                    MeasurementPlanning = 'Y',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //更新此量測項目原本

                        #endregion

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

        #region //AddQcGoodsReceiptLog -- 新增進貨檢驗作業資料 -- Ann 2025-02-07
        public string AddQcGoodsReceiptLog(int QcGoodsReceiptId, int ReceiptQty, int AcceptQty, int ReturnQty
                    , string AcceptanceDate, string QcStatus, string QuickStatus, string Remark, string QcGoodsReceiptLogFile)
        {
            try
            {
                int rowsAffected = 0;
                ErpHelper erpHelper = new ErpHelper();

                using (TransactionScope transactionScope = new TransactionScope())
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
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        string companyNo = "";
                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨檢驗單據資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrDetailId 
                                    , b.CheckQcMeasureData
                                    , c.ConfirmStatus DetailConfirmStatus, c.ReceiptQty, c.AcceptQty, c.ReturnQty, c.OrigUnitPrice
                                    , d.GrId, d.ConfirmStatus
                                    FROM MES.QcGoodsReceipt a 
                                    INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                    INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                                    INNER JOIN SCM.GoodsReceipt d ON c.GrId = d.GrId
                                    WHERE a.QcGoodsReceiptId = @QcGoodsReceiptId";
                            dynamicParameters.Add("QcGoodsReceiptId", QcGoodsReceiptId);

                            var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptResult.Count() <= 0) throw new SystemException("進貨檢驗單據錯誤!!");

                            int GrDetailId = -1;
                            int GrId = -1;
                            double OrigUnitPrice = -1;
                            foreach (var item in QcGoodsReceiptResult)
                            {
                                if (QuickStatus == "N" && item.CheckQcMeasureData != "Y" && item.CheckQcMeasureData != "P")
                                {
                                    throw new SystemException("單據狀態錯誤，需先上傳量測數據後才能進行檢驗作業!!");
                                }

                                if (item.DetailConfirmStatus != "N" || item.ConfirmStatus != "N")
                                {
                                    throw new SystemException("進貨單頭或單身狀態已核單，無法修改!!");
                                }

                                if (ReceiptQty != item.ReceiptQty)
                                {
                                    throw new SystemException("【檢驗單進貨數量】與【進貨單進貨數量】不同!!");
                                }

                                if (AcceptQty + ReturnQty != ReceiptQty)
                                {
                                    throw new SystemException("【驗收數量】+【驗退數量】需等於【進貨數量】!!");
                                }

                                GrDetailId = item.GrDetailId;
                                GrId = item.GrId;
                                OrigUnitPrice = item.OrigUnitPrice;
                            }
                            #endregion

                            #region //確認此進貨檢驗單是否已有檢驗作業紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.QcGoodsReceiptLog a 
                                    WHERE a.QcGoodsReceiptId = @QcGoodsReceiptId";
                            dynamicParameters.Add("QcGoodsReceiptId", QcGoodsReceiptId);

                            var QcGoodsReceiptLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptLogResult.Count() > 0) throw new SystemException("此進貨檢驗單是否已有檢驗作業紀錄!!");
                            #endregion

                            #region //UPDATE 進貨單相關資料

                            #region //UPDATE SCM.GrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GrDetail SET
                                    AcceptanceDate = @AcceptanceDate,
                                    AcceptQty = @AcceptQty,
                                    AvailableQty = @AcceptQty,
                                    ReturnQty = @ReturnQty,
                                    QcStatus = @QcStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                AcceptanceDate,
                                AcceptQty,
                                ReturnQty,
                                QcStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrDetailId
                            });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(erpHelper.UpdateGrTotal(GrId, -1, sqlConnection, sqlConnection2, CurrentUser));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
                            }
                            #endregion

                            #region //拋轉ERP
                            JObject updateTransferGoodsReceiptResult = JObject.Parse(erpHelper.UpdateTransferGoodsReceipt(GrId, sqlConnection, sqlConnection2, CurrentUser, CurrentCompany));
                            if (updateTransferGoodsReceiptResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateTransferGoodsReceiptResult["msg"].ToString());
                            }
                            #endregion

                            #endregion

                            #region //確認是否為急料，若為急料則由品保直接核單
                            if (QuickStatus == "Y")
                            {
                                #region //確認是否有附件
                                if (QcGoodsReceiptLogFile.Length <= 0)
                                {
                                    throw new SystemException("若為急料，需檢附證明文件!!");
                                }
                                #endregion

                                JObject confirmGrDetailResult = JObject.Parse(erpHelper.ConfirmGrDetail(GrDetailId, sqlConnection, sqlConnection2, CurrentUser));
                                if (confirmGrDetailResult["status"].ToString() != "success")
                                {
                                    throw new SystemException(confirmGrDetailResult["msg"].ToString());
                                }
                            }
                            #endregion

                            #region //INSERT MES.QcGoodsReceiptLog
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcGoodsReceiptLog (QcGoodsReceiptId, ReceiptQty, AcceptQty, ReturnQty, AcceptanceDate
                                    , QcStatus, QuickStatus, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LogId
                                    VALUES (@QcGoodsReceiptId, @ReceiptQty, @AcceptQty, @ReturnQty, @AcceptanceDate
                                    , @QcStatus, @QuickStatus, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcGoodsReceiptId,
                                    ReceiptQty,
                                    AcceptQty,
                                    ReturnQty,
                                    AcceptanceDate,
                                    QcStatus,
                                    QuickStatus,
                                    Remark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int LogId = -1;
                            foreach (var item in insertResult)
                            {
                                LogId = item.LogId;
                            }

                            rowsAffected += insertResult.Count();
                            #endregion

                            #region //INSERT MES.QcGoodsReceiptLogFile
                            if (QcGoodsReceiptLogFile.Length > 0)
                            {
                                var qcGoodsReceiptLogFileList = QcGoodsReceiptLogFile.Split(',');

                                foreach (var fileId in qcGoodsReceiptLogFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcGoodsReceiptLogFile (LogId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@LogId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LogId,
                                            FileId = fileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                }
                            }
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

        #region //AddBatchQcRecord -- 批量新增量測單據 -- Ann 2025-04-29
        public string AddBatchQcRecord(List<QcExcelFormat> qcExcelFormats)
        {
            try
            {
                string ErpDbName = "";
                int ModeId = -1;
                int? MoId = null;
                string WoErpFullNo = "";
                int MtlItemId = -1;
                string MtlItemNo = "";
                string MtlItemName = "";
                string WoErpPrefix = "";
                string WoErpNo = "";
                int QcTypeId = -1;
                string QcTypeNo = "";
                string QcTypeName = "";
                string SupportProcessFlag = "";
                string DefaultSpreadsheetData = "";
                int rowsAffected = 0;
                List<int> qcRecordIds = new List<int>();
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in DefaultFileIdResult)
                            {
                                DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                            }
                            #endregion

                            foreach (var qcExcel in qcExcelFormats)
                            {
                                if (qcExcel.QcType.Length <= 0) throw new SystemException("【量測類型】不能為空!");
                                if (qcExcel.WoErpFullNo == null) qcExcel.WoErpFullNo = "";
                                if (qcExcel.MtlItemNo == null) qcExcel.MtlItemNo = "";

                                if (qcExcel.QcType != "OQC" && qcExcel.QcType != "OIQC")
                                {
                                    throw new SystemException("此功能目前只支援【出貨檢驗】及【委外檢驗】!");
                                }
                                else if (qcExcel.QcType == "OQC" && qcExcel.WoErpFullNo.Length <= 0)
                                {
                                    throw new SystemException("【出貨檢驗】類別需要維護製令!");
                                }
                                else if (qcExcel.QcType == "OQC" && qcExcel.WoErpFullNo.Length > 0)
                                {
                                    if (qcExcel.WoErpFullNo.IndexOf("(") == -1)
                                    {
                                        qcExcel.WoErpFullNo = qcExcel.WoErpFullNo + "(1)";
                                    }
                                }
                                else if (qcExcel.QcType == "OIQC" && qcExcel.MtlItemNo.Length <= 0)
                                {
                                    throw new SystemException("【委外檢驗】類別需要維護品號!");
                                }

                                if (qcExcel.InputType.Length <= 0) throw new SystemException("【輸入方式】不能為空!");

                                #region //轉換輸入方式
                                switch (qcExcel.InputType.ToUpper())
                                {
                                    case "LETTERING":
                                        qcExcel.InputType = "Lettering";
                                        break;
                                    case "CAVITY":
                                        qcExcel.InputType = "Cavity";
                                        break;
                                    case "BARCODENO":
                                        qcExcel.InputType = "BarcodeNo";
                                        break;
                                    case "LOTNUMBER":
                                        qcExcel.InputType = "LotNumber";
                                        break;
                                    default:
                                        throw new SystemException("【輸入方式: " + qcExcel.InputType + "】錯誤!");
                                }
                                #endregion

                                if (qcExcel.QcType == "OQC")
                                {
                                    #region //判斷MES製令資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MoId, a.ModeId
                                            , (b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(varchar(10), a.WoSeq) + ')') WoErpFullNo
                                            , b.WoErpPrefix, b.WoErpNo
                                            , c.MtlItemId, c.MtlItemNo, c.MtlItemName
                                            FROM MES.ManufactureOrder a 
                                            INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                            WHERE b.WoErpPrefix + '-' + b.WoErpNo +  '(' + CONVERT(VARCHAR(10), a.WoSeq) + ')' = @WoErpFullNo";
                                    dynamicParameters.Add("WoErpFullNo", qcExcel.WoErpFullNo);

                                    var result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("製令(" + qcExcel.WoErpFullNo + ")資料錯誤!");

                                    foreach (var item in result)
                                    {
                                        MoId = item.MoId;
                                        ModeId = item.ModeId;
                                        WoErpFullNo = item.WoErpFullNo;
                                        MtlItemId = item.MtlItemId;
                                        MtlItemNo = item.MtlItemNo;
                                        MtlItemName = item.MtlItemName;
                                        WoErpPrefix = item.WoErpPrefix;
                                        WoErpNo = item.WoErpNo;
                                    }
                                    #endregion

                                    #region //判斷ERP製令狀態是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TA011, TA013
                                        FROM MOCTA
                                        WHERE TA001 = @TA001
                                        AND TA002 = @TA002";
                                    dynamicParameters.Add("TA001", WoErpPrefix);
                                    dynamicParameters.Add("TA002", WoErpNo);

                                    var MOCTAResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (MOCTAResult.Count() <= 0) throw new SystemException("製令(" + qcExcel.WoErpFullNo + ")資料錯誤!!");

                                    foreach (var item in MOCTAResult)
                                    {
                                        switch (item.TA011)
                                        {
                                            case "Y":
                                                throw new SystemException("ERP製令(" + qcExcel.WoErpFullNo + ")狀態已完工，無法開立!!");
                                            case "y":
                                                throw new SystemException("ERP製令(" + qcExcel.WoErpFullNo + ")狀態已指定完工，無法開立!!");
                                            case "V":
                                                throw new SystemException("ERP製令(" + qcExcel.WoErpFullNo + ")狀態已作廢，無法開立!!");
                                        }
                                    }
                                    #endregion
                                }
                                else if (qcExcel.QcType == "OIQC")
                                {
                                    qcExcel.InputType = "Cavity";

                                    #region //判斷品號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlItemId
                                            FROM PDM.MtlItem a 
                                            WHERE a.MtlItemNo = @MtlItemNo
                                            AND a.CompanyId = @CompanyId";
                                    dynamicParameters.Add("MtlItemNo", qcExcel.MtlItemNo);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MtlItemResult.Count() <= 0) throw new SystemException("【品號:" + qcExcel.MtlItemNo + "】資料錯誤!");

                                    MtlItemId = MtlItemResult.FirstOrDefault().MtlItemId;
                                    #endregion
                                }

                                #region //判斷量測類型資料是否正確
                                if (qcExcel.QcType == "OQC")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.QcTypeId, a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                        FROM QMS.QcType a 
                                        WHERE a.ModeId = @ModeId
                                        AND a.QcTypeNo = @QcTypeNo";
                                    dynamicParameters.Add("ModeId", ModeId);
                                    dynamicParameters.Add("QcTypeNo", qcExcel.QcType);

                                    var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                                    foreach (var item in QcTypeResult)
                                    {
                                        QcTypeId = item.QcTypeId;
                                        QcTypeNo = item.QcTypeNo;
                                        QcTypeName = item.QcTypeName;
                                        SupportProcessFlag = item.SupportProcessFlag;
                                    }
                                }
                                else
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.QcTypeId, a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                            FROM QMS.QcType a 
                                            WHERE a.QcTypeNo = @QcTypeNo";
                                    dynamicParameters.Add("QcTypeNo", qcExcel.QcType);

                                    var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                                    foreach (var item in QcTypeResult)
                                    {
                                        QcTypeId = item.QcTypeId;
                                        QcTypeNo = item.QcTypeNo;
                                        QcTypeName = item.QcTypeName;
                                        SupportProcessFlag = item.SupportProcessFlag;
                                    }
                                }
                                #endregion

                                #region //建立CheckQcMeasureData
                                string CheckQcMeasureData = "N";
                                #endregion

                                #region //INSERT MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, MoId, MtlItemId, Remark, DefaultFileId, DefaultSpreadsheetData, CheckQcMeasureData, SupportAqFlag, SupportProcessFlag
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcRecordId
                                        VALUES (@QcTypeId, @InputType, @MoId, @MtlItemId, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @CheckQcMeasureData, @SupportAqFlag, @SupportProcessFlag
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcTypeId,
                                        qcExcel.InputType,
                                        MoId,
                                        MtlItemId,
                                        qcExcel.Remark,
                                        DefaultFileId = -1,
                                        DefaultSpreadsheetData,
                                        CheckQcMeasureData,
                                        SupportAqFlag = "Y",
                                        SupportProcessFlag = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int QcRecordId = -1;
                                foreach (var item in insertResult)
                                {
                                    QcRecordId = item.QcRecordId;
                                    qcRecordIds.Add(QcRecordId);
                                }
                                #endregion

                                #region //建立量測單據Spreadsheet DATA
                                #region //取得品號允收標準
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                        , a.QcItemDesc, a.SortNumber, a.QmmDetailId, a.BallMark, a.Unit, a.QmmDetailId
                                        , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                        , c.MachineNumber
                                        , d.MachineDesc
                                        FROM PDM.MtlQcItem a
                                        INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                        LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                        LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                        WHERE a.MtlItemId = @MtlItemId
                                        AND a.[Status] = 'A' 
                                        ORDER BY a.SortNumber";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //初始化Data
                                List<Data> datas = new List<Data>();
                                Data data = new Data();
                                data = new Data()
                                {
                                    cell = "A1",
                                    css = "imported_class1",
                                    format = "text",
                                    value = "序號",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B1",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "球標",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C1",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "檢測項目",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D1",
                                    css = "imported_class3",
                                    format = "text",
                                    value = "檢測備註",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E1",
                                    css = "imported_class4",
                                    format = "text",
                                    value = "量測設備",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F1",
                                    css = "imported_class5",
                                    format = "text",
                                    value = "設計值",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G1",
                                    css = "imported_class6",
                                    format = "text",
                                    value = "上限",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "H1",
                                    css = "imported_class7",
                                    format = "text",
                                    value = "下限",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "I1",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "單位",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "J1",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "Z軸",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "K1",
                                    css = "imported_class9",
                                    format = "text",
                                    value = "量測人員",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "L1",
                                    css = "imported_class10",
                                    format = "text",
                                    value = "量測工時",
                                };
                                datas.Add(data);
                                #endregion

                                #region //設定單身量測標準
                                int row = 2;
                                foreach (var item2 in MtlQcItemResult)
                                {
                                    #region //若有機台，整理序號格式
                                    string QcItemNo = item2.QcItemNo;
                                    if (item2.MachineNumber != null)
                                    {
                                        QcItemNo = item2.QcItemNo;
                                        string firstPart = QcItemNo.Substring(0, 3);
                                        string secondPart = QcItemNo.Substring(3, 4);
                                        QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                    }
                                    #endregion

                                    string QcItemName = item2.QcItemName;

                                    #region //設定量測項目、備註、上下限
                                    data = new Data()
                                    {
                                        cell = "A" + row,
                                        css = "",
                                        format = "common",
                                        value = QcItemNo,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "B" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.BallMark != null ? item2.BallMark.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "C" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.QcItemName,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "D" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.QcItemDesc,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "E" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "F" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "G" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "H" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "I" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.Unit != null ? item2.Unit.ToString() : "",
                                    };
                                    datas.Add(data);

                                    row++;
                                    #endregion
                                }
                                #endregion

                                #region //整合Spreadsheet格式
                                List<Sheets> sheetss = new List<Sheets>();

                                Sheets sheets = new Sheets()
                                {
                                    name = "sheet1",
                                    data = datas
                                };
                                sheetss.Add(sheets);

                                #region //更新至QcRecord SpreadsheetData
                                SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                {
                                    sheets = sheetss
                                };

                                string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion
                                #endregion
                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                qcRecordIds = String.Join(",", qcRecordIds)
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

        #region //AddQcMeasurePointData -- 新增量測點資料 -- WuTc 2024-12-16
        public string AddQcMeasurePointData(List<QcMeasurePointData> PointDataSetList)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int rowsAffected = 0;
                    PointDataSetList
                        .ToList()
                        .ForEach(x =>
                        {
                            x.CreateDate = CreateDate;
                            x.LastModifiedDate = LastModifiedDate;
                            x.CreateBy = CreateBy;
                            x.LastModifiedBy = LastModifiedBy;
                        });

                    dynamicParameters = new DynamicParameters();
                    sql = @"INSERT INTO QMS.QcMeasurePointData (QcRecordFileId, LotNumber, BarcodeId, Point, PointValue, Axis, Cavity, PointSite, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@QcRecordFileId, @LotNumber, @BarcodeId, @Point, @PointValue, @Axis, @Cavity, @PointSite, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                    rowsAffected += sqlConnection.Execute(sql, PointDataSetList);

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

        #region //AddOutsourcingQcRecord -- 新增量測紀錄單頭  -- GPAI 2025-04-10
        public string AddOutsourcingQcRecord(int QcNoticeId, int QcTypeId, int MtlItemId, int QmmDetailId, int QcItemId, string Remark, int CurrentFileId, string SupportAqFlag, string LotNumber
            , string QcRecordFile, string InputType, string ServerPath, string ServerPath2, string ResolveFile, string ResolveFileJson, string QcRecordFileByNas, string QcMeasureDataJson)
        {
            try
            {
                string ErpDbName = "";
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (InputType.Length < 0) throw new SystemException("【輸入方式】不能為空!");
                            //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Cavity") throw new SystemException("【穴號模式】只能由全吋檢或成型工程檢使用!!");
                            //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Picture") throw new SystemException("【面型圖模式】只能由全吋檢或成型工程檢使用!!");

                            #region //判斷MES製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * 
                                    FROM PDM.MtlItem 
                                    WHERE MtlItemId  = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("MES製令資料錯誤!");

                            string MtlItemNo = "";
                            string MtlItemName = "";
                            foreach (var item in result)
                            {
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                            }
                            #endregion

                            #region //判斷ERP製令狀態是否正確
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT TA011, TA013
                            //        FROM MOCTA
                            //        WHERE TA001 = @TA001
                            //        AND TA002 = @TA002";
                            //dynamicParameters.Add("TA001", WoErpPrefix);
                            //dynamicParameters.Add("TA002", WoErpNo);

                            //var MOCTAResult = sqlConnection2.Query(sql, dynamicParameters);

                            //if (MOCTAResult.Count() <= 0) throw new SystemException("ERP製令資料錯誤!!");

                            //foreach (var item in MOCTAResult)
                            //{
                            //    if (SupportAqFlag == "Y" && (item.TA011 == "Y" || item.TA011 == "y" || item.TA013 == "V")) throw new SystemException("ERP製令狀態無法開立!!");
                            //}
                            #endregion

                            #region //判斷量測類型資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                    FROM QMS.QcType a 
                                    WHERE a.QcTypeId = @QcTypeId";
                            dynamicParameters.Add("QcTypeId", QcTypeId);

                            var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                            string QcTypeNo = "";
                            string QcTypeName = "";
                            string SupportProcessFlag = "";
                            foreach (var item in QcTypeResult)
                            {
                                if (item.SupportProcessFlag == "Y" /*&& MoProcessId <= 0*/) throw new SystemException("【量測製程】不能為空!");
                                QcTypeNo = item.QcTypeNo;
                                QcTypeName = item.QcTypeName;
                                SupportProcessFlag = item.SupportProcessFlag;
                            }
                            #endregion

                            //if (MoProcessId > 0)
                            //{
                            //    #region //判斷製程資料是否正確
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"SELECT TOP 1 1
                            //            FROM MES.MoProcess
                            //            WHERE MoProcessId = @MoProcessId";
                            //    dynamicParameters.Add("MoProcessId", MoProcessId);

                            //    var MoProcessResult = sqlConnection.Query(sql, dynamicParameters);
                            //    if (MoProcessResult.Count() <= 0) throw new SystemException("製令製程資料錯誤!");
                            //    #endregion
                            //}

                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            string DefaultSpreadsheetData = "{\"sheets\": [{\"name\": \"sheet1\", \"data\": [{\"cell\": \"A1\", \"css\": \"imported_class1\", \"format\": \"text\", \"value\": \"序號\"}, {\"cell\": \"B1\", \"css\": \"imported_class2\", \"format\": \"text\", \"value\": \"球標\"}, {\"cell\": \"C1\", \"css\": \"imported_class2\", \"format\": \"text\", \"value\": \"檢測項目\"}, {\"cell\": \"D1\", \"css\": \"imported_class3\", \"format\": \"text\", \"value\": \"檢測備註\"}, {\"cell\": \"E1\", \"css\": \"imported_class4\", \"format\": \"text\", \"value\": \"量測設備\"}, {\"cell\": \"F1\", \"css\": \"imported_class5\", \"format\": \"text\", \"value\": \"設計值\"}, {\"cell\": \"G1\", \"css\": \"imported_class6\", \"format\": \"text\", \"value\": \"上限\"}, {\"cell\": \"H1\", \"css\": \"imported_class7\", \"format\": \"text\", \"value\": \"下限\"}, {\"cell\": \"I1\", \"css\": \"imported_class8\", \"format\": \"text\", \"value\": \"單位\"}, {\"cell\": \"J1\", \"css\": \"imported_class8\", \"format\": \"text\", \"value\": \"Z軸\"}, {\"cell\": \"K1\", \"css\": \"imported_class9\", \"format\": \"text\", \"value\": \"量測人員\"}, {\"cell\": \"L1\", \"css\": \"imported_class10\", \"format\": \"text\", \"value\": \"量測工時\"}, {\"cell\": \"A2\", \"format\": \"common\", \"value\": \"D02z101\"}, {\"cell\": \"C2\", \"format\": \"common\", \"value\": \"判定量測單01\"}, {\"cell\": \"D2\", \"format\": \"common\", \"value\": \"判定量測單01\"}, {\"cell\": \"F2\", \"format\": \"common\", \"value\": 1}, {\"cell\": \"G2\", \"format\": \"common\", \"value\": 0}, {\"cell\": \"H2\", \"format\": \"common\", \"value\": 0}], \"cols\": [{\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}, {\"width\": 120}], \"rows\": [{\"height\": 32}, {\"height\": 32}]}], \"styles\": {\"imported_class1\": null, \"imported_class2\": null, \"imported_class3\": null, \"imported_class4\": null, \"imported_class5\": null, \"imported_class6\": null, \"imported_class7\": null, \"imported_class8\": null, \"imported_class9\": null, \"imported_class10\": null}, \"formats\": [{\"name\": \"Common\", \"id\": \"common\", \"mask\": \"\", \"example\": \"1500.31\"}, {\"name\": \"Number\", \"id\": \"number\", \"mask\": \"#,##0.00\", \"example\": \"1500.31\"}, {\"name\": \"Percent\", \"id\": \"percent\", \"mask\": \"#,##0.00%\", \"example\": \"15.0031\"}, {\"name\": \"Currency\", \"id\": \"currency\", \"mask\": \"$#,##0.00\", \"example\": \"1500.31\"}, {\"name\": \"Date\", \"id\": \"date\", \"mask\": \"mm-dd-yy\", \"example\": \"44490\", \"dateFormat\": \"%d/%m/%Y\"}, {\"name\": \"Time\", \"id\": \"time\", \"mask\": \"h:mm:ss am/pm\", \"example\": \"0.5625\", \"timeFormat\": 12}, {\"name\": \"Text\", \"id\": \"text\", \"mask\": \"@\", \"example\": \"some text\"}]}";
                            
                            #endregion

                            #region //建立CheckQcMeasureData
                            string CheckQcMeasureData = "N";
                            //if (InputType == "Cavity" || InputType == "Picture")
                            //{
                            //    if (QcRecordFile.Length > 0)
                            //    {
                            //        CheckQcMeasureData = "Y";
                            //    }
                            //}
                            #endregion

                            #region //INSERT MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcRecord (QcNoticeId, QcTypeId, InputType, MtlItemId, Remark, DefaultFileId, CurrentFileId, DefaultSpreadsheetData, CheckQcMeasureData, SupportAqFlag, SupportProcessFlag, ResolveFileJson
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcRecordId
                                    VALUES (@QcNoticeId, @QcTypeId, @InputType, @MtlItemId, @Remark, @DefaultFileId, @CurrentFileId, @DefaultSpreadsheetData, @CheckQcMeasureData, @SupportAqFlag, @SupportProcessFlag, @ResolveFileJson
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId = QcNoticeId <= 0 ? (int?)null : QcNoticeId,
                                    QcTypeId,
                                    InputType,
                                    MtlItemId,
                                    //MoId,
                                    //MoProcessId = MoProcessId > 0 ? MoProcessId : (int?)null,
                                    Remark,
                                    DefaultFileId = -1,
                                    CurrentFileId,
                                    DefaultSpreadsheetData,
                                    CheckQcMeasureData,
                                    SupportAqFlag,
                                    SupportProcessFlag,
                                    ResolveFileJson,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int QcRecordId = -1;
                            foreach (var item in insertResult)
                            {
                                QcRecordId = item.QcRecordId;
                            }
                            #endregion

                            #region //INSERT MES.QcRecordFile
                            if (QcRecordFile.Length > 0)
                            {
                                var qcRecordFileList = QcRecordFile.Split(',');

                                foreach (var qcRecordFileId in qcRecordFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileId = qcRecordFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT MES.QcRecordFile By NAS
                            if (QcRecordFileByNas.Length > 0)
                            {
                                JObject uploadFileJson = JObject.Parse(QcRecordFileByNas);

                                foreach (var item in uploadFileJson["uploadFileInfo"])
                                {
                                    string fileType = item["FileType"]?.ToString() ?? "other";
                                    string inputType = item["InputType"]?.ToString() ?? null;
                                    string barcodeId = item["BarcodeId"]?.ToString() ?? null;

                                    //SplitCoating (SC) 分光鍍膜 file-management
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, InputType, BarcodeId, PhysicalPath, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @InputType, @BarcodeId, @PhysicalPath, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = fileType,
                                            InputType = inputType,
                                            BarcodeId = barcodeId,
                                            PhysicalPath = item["FilePath"].ToString(),
                                            LotNumber = LotNumber ?? null,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT QMS.QcMeasureData
                            if (QcMeasureDataJson.Length > 0)
                            {
                                if (string.IsNullOrWhiteSpace(LotNumber)) throw new SystemException("【批號】不能為空!");

                                #region //取得量測項目
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QcItemId, a.QcItemNo, a.QcItemName, a.QcType, ISNULL(b.QcItemDesc, a.QcItemDesc) QcItemDesc, b.*
                                        FROM QMS.QcItem a
                                        OUTER APPLY (
	                                        SELECT aa.QcItemDesc, aa.DesignValue, aa.UpperTolerance, aa.LowerTolerance, aa.SortNumber
	                                        , ab.QmmDetailId, ab.MachineNumber
	                                        , ac.MachineNo, ac.MachineName, ad.ShopName 
	                                        FROM PDM.MtlQcItem aa
	                                        INNER JOIN QMS.QmmDetail ab ON ab.QmmDetailId = aa.QmmDetailId
	                                        INNER JOIN MES.Machine ac ON ac.MachineId = ab.MachineId
	                                        INNER JOIN MES.WorkShop ad ON ad.ShopId = ac.ShopId AND ad.CompanyId = @CompanyId
	                                        WHERE aa.QcItemId = a.QcItemId
	                                        AND ad.CompanyId = @CompanyId
	                                        AND aa.MtlItemId = @MtlItemId
                                        ) b
                                        WHERE a.QcItemId = @QcItemId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("QcItemId", QcItemId);
                                var resultQcItem = sqlConnection.Query(sql, dynamicParameters);
                                if (!resultQcItem.Any()) throw new SystemException("【量測項目】資料錯誤!");
                                string QcItemDesc = string.Empty;
                                string DesignValue = string.Empty;
                                string UpperTolerance = string.Empty;
                                string LowerTolerance = string.Empty;
                                foreach (var item in resultQcItem)
                                {
                                    QcItemDesc = item.QcItemDesc;
                                    DesignValue = item.DesignValue ?? "OK";
                                    UpperTolerance = item.UpperTolerance;
                                    LowerTolerance = item.LowerTolerance;
                                    QmmDetailId = item.QmmDetailId ?? QmmDetailId;
                                }
                                #endregion

                                #region //取得批號 (停用)
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT a.LotNumberId, a.MtlItemId, a.LotNumberNo
                                //        FROM SCM.LotNumber a
                                //        WHERE a.MtlItemId = @MtlItemId";
                                //dynamicParameters.Add("MtlItemId", MtlItemId);
                                //var resultLotNumber = sqlConnection.Query(sql, dynamicParameters);
                                //string LotNumber = string.Empty;
                                //foreach (var item in resultLotNumber)
                                //{
                                //    LotNumber = item.LotNumberNo;
                                //}
                                #endregion

                                JObject qcMeasureData = JObject.Parse(QcMeasureDataJson);

                                foreach (var item in qcMeasureData["data"])
                                {
                                    int.TryParse(item["BarcodeId"]?.ToString(), out int BarcodeId);
                                    string MeasureValue = item["MeasureValue"]?.ToString() ?? "";

                                    #region //新增量測資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO QMS.QcMeasureData (QcRecordId, QcItemId, QcItemDesc, DesignValue, BarcodeId, QmmDetailId, MeasureValue, QcResult, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QmdId
                                            VALUES (@QcRecordId, @QcItemId, @QcItemDesc, @DesignValue, @BarcodeId, @QmmDetailId, @MeasureValue, @QcResult, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            QcItemId,
                                            QcItemDesc,
                                            DesignValue,
                                            BarcodeId,
                                            QmmDetailId,
                                            MeasureValue,
                                            QcResult = MeasureValue == "OK" ? "P" : "F",
                                            LotNumber = LotNumber ?? null,
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
                            #endregion

                            #region //INSERT ResolveFile
                            if (ResolveFile.Length > 0)
                            {
                                var resolveFileList = ResolveFile.Split(',');

                                foreach (var resolveFileId in resolveFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = "resolve",
                                            FileId = resolveFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //建立量測單據Spreadsheet DATA
                            #region //取得品號允收標準
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                    , a.QcItemDesc, a.SortNumber, a.QmmDetailId, a.BallMark, a.Unit, a.QmmDetailId
                                    , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                    , c.MachineNumber
                                    , d.MachineDesc
                                    FROM PDM.MtlQcItem a
                                    INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                    LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                    LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.[Status] = 'A' 
                                    ORDER BY a.SortNumber";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //初始化Data
                            List<Data> datas = new List<Data>();
                            Data data = new Data();
                            data = new Data()
                            {
                                cell = "A1",
                                css = "imported_class1",
                                format = "text",
                                value = "序號",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "B1",
                                css = "imported_class2",
                                format = "text",
                                value = "球標",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "C1",
                                css = "imported_class2",
                                format = "text",
                                value = "檢測項目",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "D1",
                                css = "imported_class3",
                                format = "text",
                                value = "檢測備註",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "E1",
                                css = "imported_class4",
                                format = "text",
                                value = "量測設備",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "F1",
                                css = "imported_class5",
                                format = "text",
                                value = "設計值",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "G1",
                                css = "imported_class6",
                                format = "text",
                                value = "上限",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "H1",
                                css = "imported_class7",
                                format = "text",
                                value = "下限",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "I1",
                                css = "imported_class8",
                                format = "text",
                                value = "單位",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "J1",
                                css = "imported_class8",
                                format = "text",
                                value = "Z軸",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "K1",
                                css = "imported_class9",
                                format = "text",
                                value = "量測人員",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "L1",
                                css = "imported_class10",
                                format = "text",
                                value = "量測工時",
                            };
                            datas.Add(data);
                            #endregion

                            #region //加入預帶項目
                            data = new Data()
                            {
                                cell = "A2",
                                css = "imported_class1",
                                format = "text",
                                value = "D02z101",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "B2",
                                css = "imported_class2",
                                format = "text",
                                value = "",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "C2",
                                css = "imported_class2",
                                format = "text",
                                value = "判定量測單01",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "D2",
                                css = "imported_class3",
                                format = "text",
                                value = "判定量測單01",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "E2",
                                css = "imported_class4",
                                format = "text",
                                value = "",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "F2",
                                css = "imported_class5",
                                format = "text",
                                value = "0",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "G2",
                                css = "imported_class6",
                                format = "text",
                                value = "1",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "H2",
                                css = "imported_class7",
                                format = "text",
                                value = "1",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "I2",
                                css = "imported_class8",
                                format = "text",
                                value = "",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "J2",
                                css = "imported_class8",
                                format = "text",
                                value = "",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "K2",
                                css = "imported_class9",
                                format = "text",
                                value = "",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "L2",
                                css = "imported_class10",
                                format = "text",
                                value = "",
                            };
                            datas.Add(data);

                            data = new Data()
                            {
                                cell = "M1",
                                css = "imported_class10",
                                format = "text",
                                value = "1-1",
                            };
                            datas.Add(data);
                            #endregion

                            #region //設定單身量測標準
                            int row = 3;
                            foreach (var item2 in MtlQcItemResult)
                            {
                                #region //若有機台，整理序號格式
                                string QcItemNo = item2.QcItemNo;
                                if (item2.MachineNumber != null)
                                {
                                    QcItemNo = item2.QcItemNo;
                                    string firstPart = QcItemNo.Substring(0, 3);
                                    string secondPart = QcItemNo.Substring(3, 4);
                                    QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                }
                                #endregion

                                string QcItemName = item2.QcItemName;

                                #region //設定量測項目、備註、上下限
                                data = new Data()
                                {
                                    cell = "A" + row,
                                    css = "",
                                    format = "common",
                                    value = QcItemNo,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.BallMark != null ? item2.BallMark.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.QcItemName,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.QcItemDesc,
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "H" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "I" + row,
                                    css = "",
                                    format = "text",
                                    value = item2.Unit != null ? item2.Unit.ToString() : "",
                                };
                                datas.Add(data);

                                row++;
                                #endregion
                            }
                            #endregion



                            #region //整合Spreadsheet格式
                            List<Sheets> sheetss = new List<Sheets>();

                            Sheets sheets = new Sheets()
                            {
                                name = "sheet1",
                                data = datas
                            };
                            sheetss.Add(sheets);

                            #region //更新至QcRecord SpreadsheetData
                            SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                            {
                                sheets = sheetss
                            };

                            string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    SpreadsheetData = @SpreadsheetData,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SpreadsheetData,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion
                            #endregion

                            #region //將量測EXCEL先存回實體路徑並重新讀取
                            string FileName = "";
                            if (CurrentFileId > 0)
                            {
                                #region //取得原本加密檔案資料，並重新上傳
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId, a.FileContent, a.FileName, a.FileExtension, a.FileSize, a.ClientIP, a.Source
                                        FROM BAS.[File] a
                                        WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", CurrentFileId);

                                var FileResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in FileResult)
                                {
                                    FileName = item.FileName;
                                    ServerPath = Path.Combine(ServerPath, item.FileName + "-" + item.FileId + item.FileExtension);
                                    byte[] fileContent = (byte[])item.FileContent;
                                    File.WriteAllBytes(ServerPath, fileContent); // Requires System.IO

                                    #region //將檔案重新上傳
                                    System.Net.WebClient wc = new System.Net.WebClient();
                                    byte[] newFilebytes = wc.DownloadData(ServerPath);

                                    dynamicParameters = new DynamicParameters();
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
                                            CompanyId = CurrentCompany,
                                            item.FileName,
                                            FileContent = newFilebytes,
                                            item.FileExtension,
                                            item.FileSize,
                                            item.ClientIP,
                                            item.Source,
                                            DeleteStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    int newFileId = -1;
                                    foreach (var item2 in insertResult)
                                    {
                                        newFileId = item2.FileId;
                                    }
                                    #endregion

                                    #region //UPDATE原本FileId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcRecord SET
                                            CurrentFileId = @CurrentFileId,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CurrentFileId = newFileId,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QcRecordId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除實體檔案
                                    File.Delete(ServerPath);
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //統整回傳前端資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId, a.CurrentFileId
                                    , (
                                        SELECT x.FileId
                                        FROM MES.QcRecordFile x
                                        WHERE x.QcRecordId = a.QcRecordId
                                        FOR JSON PATH, ROOT('data')
                                    ) QcRecordFile
                                    FROM MES.QcRecord a
                                    WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            if ((CheckQcMeasureData == "Y" || CheckQcMeasureData == "P") && CurrentFileId > 0 && (QcTypeNo == "PVTQC" || QcTypeNo == "TQC"))
                            {
                                SendQcMail(sqlConnection, QcRecordId, "", MtlItemNo, MtlItemName, QcTypeNo, QcTypeName
                                    , Remark, FileName, ServerPath2, -1);
                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                data = QcRecordResult
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

        #region //AddBatchOutsourcingQcRecord -- 批量新增量測單據 -- GPAI 2025-05-06
        public string AddBatchOutsourcingQcRecord(List<QcExcelFormat> qcExcelFormats)
        {
            try
            {
                string ErpDbName = "";
                //int ModeId = -1;
                //int? MoId = null;
               // string WoErpFullNo = "";
                int MtlItemId = -1;
                string MtlItemNo = "";
                string MtlItemName = "";
               // string WoErpPrefix = "";
               // string WoErpNo = "";
                int QcTypeId = -1;
                string QcTypeNo = "";
                string QcTypeName = "";
                string SupportProcessFlag = "";
                string DefaultSpreadsheetData = "";
                int rowsAffected = 0;
                List<int> qcRecordIds = new List<int>();
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in DefaultFileIdResult)
                            {
                                DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                            }
                            #endregion

                            foreach (var qcExcel in qcExcelFormats)
                            {
                                if (qcExcel.QcType.Length <= 0) throw new SystemException("【量測類型】不能為空!");
                                if (qcExcel.WoErpFullNo == null) qcExcel.WoErpFullNo = "";
                                if (qcExcel.MtlItemNo == null) qcExcel.MtlItemNo = "";

                                if (qcExcel.QcType != "OQC" && qcExcel.QcType != "OIQC")
                                {
                                    throw new SystemException("此功能目前只支援【出貨檢驗】及【委外檢驗】!");
                                }
                                else if (qcExcel.QcType == "OQC" && qcExcel.WoErpFullNo.Length <= 0)
                                {
                                    throw new SystemException("【出貨檢驗】類別需要維護製令!");
                                }
                                else if (qcExcel.QcType == "OQC" && qcExcel.WoErpFullNo.Length > 0)
                                {
                                    if (qcExcel.WoErpFullNo.IndexOf("(") == -1)
                                    {
                                        qcExcel.WoErpFullNo = qcExcel.WoErpFullNo + "(1)";
                                    }
                                }
                                else if (qcExcel.QcType == "OIQC" && qcExcel.MtlItemNo.Length <= 0)
                                {
                                    throw new SystemException("【委外檢驗】類別需要維護品號!");
                                }

                                if (qcExcel.InputType.Length <= 0) throw new SystemException("【輸入方式】不能為空!");

                                #region //轉換輸入方式
                                switch (qcExcel.InputType.ToUpper())
                                {
                                    case "LETTERING":
                                        qcExcel.InputType = "Lettering";
                                        break;
                                    case "CAVITY":
                                        qcExcel.InputType = "Cavity";
                                        break;
                                    case "BARCODENO":
                                        qcExcel.InputType = "BarcodeNo";
                                        break;
                                    case "LOTNUMBER":
                                        qcExcel.InputType = "LotNumber";
                                        break;
                                    default:
                                        throw new SystemException("【輸入方式: " + qcExcel.InputType + "】錯誤!");
                                }
                                #endregion

                                if (qcExcel.QcType != "OIQC")
                                {
                                    throw new SystemException("量測類型非委託量測!!");
                                }
                                else if (qcExcel.QcType == "OIQC")
                                {
                                    qcExcel.InputType = "Cavity";

                                    #region //判斷品號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlItemId
                                            FROM PDM.MtlItem a 
                                            WHERE a.MtlItemNo = @MtlItemNo
                                            AND a.CompanyId = @CompanyId";
                                    dynamicParameters.Add("MtlItemNo", qcExcel.MtlItemNo);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MtlItemResult.Count() <= 0) throw new SystemException("【品號:" + qcExcel.MtlItemNo + "】資料錯誤!");

                                    MtlItemId = MtlItemResult.FirstOrDefault().MtlItemId;
                                    #endregion
                                }

                                #region //判斷量測類型資料是否正確
                                if (qcExcel.QcType != "OIQC")
                                {
                                    throw new SystemException("量測類型非委託量測!!");
                                }
                                else
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.QcTypeId, a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                            FROM QMS.QcType a 
                                            WHERE a.QcTypeNo = @QcTypeNo";
                                    dynamicParameters.Add("QcTypeNo", qcExcel.QcType);

                                    var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                                    foreach (var item in QcTypeResult)
                                    {
                                        QcTypeId = item.QcTypeId;
                                        QcTypeNo = item.QcTypeNo;
                                        QcTypeName = item.QcTypeName;
                                        SupportProcessFlag = item.SupportProcessFlag;
                                    }
                                }
                                #endregion

                                #region //建立CheckQcMeasureData
                                string CheckQcMeasureData = "N";
                                #endregion

                                #region //INSERT MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, MtlItemId, Remark, DefaultFileId, DefaultSpreadsheetData, CheckQcMeasureData, SupportAqFlag, SupportProcessFlag
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcRecordId
                                        VALUES (@QcTypeId, @InputType, @MtlItemId, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @CheckQcMeasureData, @SupportAqFlag, @SupportProcessFlag
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcTypeId,
                                        qcExcel.InputType,
                                        MtlItemId,
                                        qcExcel.Remark,
                                        DefaultFileId = -1,
                                        DefaultSpreadsheetData,
                                        CheckQcMeasureData,
                                        SupportAqFlag = "N",
                                        SupportProcessFlag = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int QcRecordId = -1;
                                foreach (var item in insertResult)
                                {
                                    QcRecordId = item.QcRecordId;
                                    qcRecordIds.Add(QcRecordId);
                                }
                                #endregion

                                #region //建立量測單據Spreadsheet DATA
                                #region //取得品號允收標準
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                        , a.QcItemDesc, a.SortNumber, a.QmmDetailId, a.BallMark, a.Unit, a.QmmDetailId
                                        , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                        , c.MachineNumber
                                        , d.MachineDesc
                                        FROM PDM.MtlQcItem a
                                        INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                        LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                        LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                        WHERE a.MtlItemId = @MtlItemId
                                        AND a.[Status] = 'A' 
                                        ORDER BY a.SortNumber";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //初始化Data
                                List<Data> datas = new List<Data>();
                                Data data = new Data();
                                data = new Data()
                                {
                                    cell = "A1",
                                    css = "imported_class1",
                                    format = "text",
                                    value = "序號",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B1",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "球標",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C1",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "檢測項目",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D1",
                                    css = "imported_class3",
                                    format = "text",
                                    value = "檢測備註",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E1",
                                    css = "imported_class4",
                                    format = "text",
                                    value = "量測設備",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F1",
                                    css = "imported_class5",
                                    format = "text",
                                    value = "設計值",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G1",
                                    css = "imported_class6",
                                    format = "text",
                                    value = "上限",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "H1",
                                    css = "imported_class7",
                                    format = "text",
                                    value = "下限",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "I1",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "單位",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "J1",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "Z軸",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "K1",
                                    css = "imported_class9",
                                    format = "text",
                                    value = "量測人員",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "L1",
                                    css = "imported_class10",
                                    format = "text",
                                    value = "量測工時",
                                };
                                datas.Add(data);
                                #endregion

                                #region //設定單身量測標準
                                int row = 2;
                                foreach (var item2 in MtlQcItemResult)
                                {
                                    #region //若有機台，整理序號格式
                                    string QcItemNo = item2.QcItemNo;
                                    if (item2.MachineNumber != null)
                                    {
                                        QcItemNo = item2.QcItemNo;
                                        string firstPart = QcItemNo.Substring(0, 3);
                                        string secondPart = QcItemNo.Substring(3, 4);
                                        QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                    }
                                    #endregion

                                    string QcItemName = item2.QcItemName;

                                    #region //設定量測項目、備註、上下限
                                    data = new Data()
                                    {
                                        cell = "A" + row,
                                        css = "",
                                        format = "common",
                                        value = QcItemNo,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "B" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.BallMark != null ? item2.BallMark.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "C" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.QcItemName,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "D" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.QcItemDesc,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "E" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "F" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "G" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "H" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "I" + row,
                                        css = "",
                                        format = "text",
                                        value = item2.Unit != null ? item2.Unit.ToString() : "",
                                    };
                                    datas.Add(data);

                                    row++;
                                    #endregion
                                }
                                #endregion

                                #region //加入預帶項目
                                data = new Data()
                                {
                                    cell = "A2",
                                    css = "imported_class1",
                                    format = "text",
                                    value = "D02z101",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B2",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C2",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "判定量測單01",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D2",
                                    css = "imported_class3",
                                    format = "text",
                                    value = "判定量測單01",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E2",
                                    css = "imported_class4",
                                    format = "text",
                                    value = "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F2",
                                    css = "imported_class5",
                                    format = "text",
                                    value = "0",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G2",
                                    css = "imported_class6",
                                    format = "text",
                                    value = "1",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "H2",
                                    css = "imported_class7",
                                    format = "text",
                                    value = "1",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "I2",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "J2",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "K2",
                                    css = "imported_class9",
                                    format = "text",
                                    value = "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "L2",
                                    css = "imported_class10",
                                    format = "text",
                                    value = "",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "M1",
                                    css = "imported_class10",
                                    format = "text",
                                    value = "1-1",
                                };
                                datas.Add(data);
                                #endregion


                                #region //整合Spreadsheet格式
                                List<Sheets> sheetss = new List<Sheets>();

                                Sheets sheets = new Sheets()
                                {
                                    name = "sheet1",
                                    data = datas
                                };
                                sheetss.Add(sheets);

                                #region //更新至QcRecord SpreadsheetData
                                SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                {
                                    sheets = sheetss
                                };

                                string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion
                                #endregion
                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                qcRecordIds = String.Join(",", qcRecordIds)
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

        #region //AddBatchQcRecordMulti -- 批量新增量測單據(多項目版本) -- Shintokoru 2025-05-20
        public string AddBatchQcRecordMulti(List<QcExcelFormatMulti> qcExcelFormats)
        {
            try
            {
                string ErpDbName = "";
                int ModeId = -1;
                int? MoId = null;
                string WoErpFullNo = "";
                int MtlItemId = -1;
                string MtlItemNo = "";
                string MtlItemName = "";
                string WoErpPrefix = "";
                string WoErpNo = "";
                int QcTypeId = -1;
                string QcTypeNo = "";
                string QcTypeName = "";
                string SupportProcessFlag = "";
                string DefaultSpreadsheetData = "";
                int rowsAffected = 0;
                List<int> qcRecordIds = new List<int>();
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in DefaultFileIdResult)
                            {
                                DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                            }
                            #endregion

                            string lastGroup = "";
                            string lastQcType = "";
                            string lastOrderRemark = "";
                            string lastWoErpFullNo = "";
                            string lastLotNumber = "";
                            string lastRemark = "";

                            foreach (var item in qcExcelFormats)
                            {
                                // Group
                                if (!string.IsNullOrWhiteSpace(item.Group))
                                    lastGroup = item.Group;
                                else
                                    item.Group = lastGroup;
                                // QcType
                                if (!string.IsNullOrWhiteSpace(item.QcType))
                                    lastQcType = item.QcType;
                                else
                                    item.QcType = lastQcType;
                                // WoErpFullNo
                                if (!string.IsNullOrWhiteSpace(item.WoErpFullNo))
                                    lastWoErpFullNo = item.WoErpFullNo;
                                else
                                    item.WoErpFullNo = lastWoErpFullNo;
                                // LotNumber
                                if (!string.IsNullOrWhiteSpace(item.LotNumber))
                                    lastLotNumber = item.LotNumber;
                                else
                                    item.LotNumber = lastLotNumber;
                                // LotNumber
                                if (!string.IsNullOrWhiteSpace(item.OrderRemark))
                                    lastOrderRemark = item.OrderRemark;
                                else
                                    item.OrderRemark = lastOrderRemark;

                                // Remark
                                if (!string.IsNullOrWhiteSpace(item.OrderRemark))
                                    lastRemark = item.OrderRemark;
                                else
                                    item.OrderRemark = lastRemark;
                            }

                            var groupedData = qcExcelFormats
                                .GroupBy(x => new { x.Group, x.QcType, x.WoErpFullNo, x.LotNumber, x.OrderRemark })
                                .ToList();

                            foreach (var group in groupedData)
                            {
                                var key = group.Key;

                                if (string.IsNullOrWhiteSpace(key.Group)) throw new SystemException("【群組】不能為空!");
                                if (string.IsNullOrWhiteSpace(key.QcType)) throw new SystemException("【量測類型】不能為空!");
                                if (string.IsNullOrWhiteSpace(key.WoErpFullNo)) throw new SystemException("【製令】不能為空!");
                                if (string.IsNullOrWhiteSpace(key.LotNumber)) throw new SystemException("【批號】不能為空!");

                                #region //判斷MES製令資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId, a.ModeId
                                            , (b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(varchar(10), a.WoSeq) + ')') WoErpFullNo
                                            , b.WoErpPrefix, b.WoErpNo
                                            , c.MtlItemId, c.MtlItemNo, c.MtlItemName
                                            FROM MES.ManufactureOrder a 
                                            INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                            WHERE b.WoErpPrefix + '-' + b.WoErpNo +  '(' + CONVERT(VARCHAR(10), a.WoSeq) + ')' = @WoErpFullNo";
                                dynamicParameters.Add("WoErpFullNo", key.WoErpFullNo);

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("製令(" + key.WoErpFullNo + ")資料錯誤!");

                                foreach (var item in result)
                                {
                                    MoId = item.MoId;
                                    ModeId = item.ModeId;
                                    WoErpFullNo = item.WoErpFullNo;
                                    MtlItemId = item.MtlItemId;
                                    MtlItemNo = item.MtlItemNo;
                                    MtlItemName = item.MtlItemName;
                                    WoErpPrefix = item.WoErpPrefix;
                                    WoErpNo = item.WoErpNo;
                                }
                                #endregion

                                #region //判斷ERP製令狀態是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TA011, TA013
                                        FROM MOCTA
                                        WHERE TA001 = @TA001
                                        AND TA002 = @TA002";
                                dynamicParameters.Add("TA001", WoErpPrefix);
                                dynamicParameters.Add("TA002", WoErpNo);

                                var MOCTAResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (MOCTAResult.Count() <= 0) throw new SystemException("製令(" + key.WoErpFullNo + ")資料錯誤!!");

                                foreach (var item in MOCTAResult)
                                {
                                    switch (item.TA011)
                                    {
                                        case "Y":
                                            throw new SystemException("ERP製令(" + key.WoErpFullNo + ")狀態已完工，無法開立!!");
                                        case "y":
                                            throw new SystemException("ERP製令(" + key.WoErpFullNo + ")狀態已指定完工，無法開立!!");
                                        case "V":
                                            throw new SystemException("ERP製令(" + key.WoErpFullNo + ")狀態已作廢，無法開立!!");
                                    }
                                }
                                #endregion

                                #region //判斷量測類型資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QcTypeId, a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                        FROM QMS.QcType a 
                                        WHERE a.ModeId = @ModeId
                                        AND a.QcTypeNo = @QcTypeNo";
                                dynamicParameters.Add("ModeId", ModeId);
                                dynamicParameters.Add("QcTypeNo", key.QcType);

                                var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                                foreach (var item in QcTypeResult)
                                {
                                    QcTypeId = item.QcTypeId;
                                    QcTypeNo = item.QcTypeNo;
                                    QcTypeName = item.QcTypeName;
                                    SupportProcessFlag = item.SupportProcessFlag;
                                }
                                #endregion

                                #region //建立CheckQcMeasureData
                                string CheckQcMeasureData = "N";
                                #endregion

                                #region //INSERT MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, MoId, MtlItemId, Remark, DefaultFileId, DefaultSpreadsheetData, CheckQcMeasureData, SupportAqFlag, SupportProcessFlag
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcRecordId
                                        VALUES (@QcTypeId, @InputType, @MoId, @MtlItemId, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @CheckQcMeasureData, @SupportAqFlag, @SupportProcessFlag
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcTypeId,
                                        InputType = "LotNumber",
                                        MoId,
                                        MtlItemId,
                                        Remark = key.OrderRemark,
                                        DefaultFileId = -1,
                                        DefaultSpreadsheetData,
                                        CheckQcMeasureData,
                                        SupportAqFlag = "N",
                                        SupportProcessFlag = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int QcRecordId = -1;
                                foreach (var item in insertResult)
                                {
                                    QcRecordId = item.QcRecordId;
                                    qcRecordIds.Add(QcRecordId);
                                }
                                #endregion

                                #region //初始化Data
                                List<Data> datas = new List<Data>();
                                Data data = new Data();
                                data = new Data()
                                {
                                    cell = "A1",
                                    css = "imported_class1",
                                    format = "text",
                                    value = "序號",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "B1",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "球標",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "C1",
                                    css = "imported_class2",
                                    format = "text",
                                    value = "檢測項目",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "D1",
                                    css = "imported_class3",
                                    format = "text",
                                    value = "檢測備註",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "E1",
                                    css = "imported_class4",
                                    format = "text",
                                    value = "量測設備",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "F1",
                                    css = "imported_class5",
                                    format = "text",
                                    value = "設計值",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "G1",
                                    css = "imported_class6",
                                    format = "text",
                                    value = "上限",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "H1",
                                    css = "imported_class7",
                                    format = "text",
                                    value = "下限",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "I1",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "單位",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "J1",
                                    css = "imported_class8",
                                    format = "text",
                                    value = "Z軸",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "K1",
                                    css = "imported_class9",
                                    format = "text",
                                    value = "量測人員",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "L1",
                                    css = "imported_class10",
                                    format = "text",
                                    value = "量測工時",
                                };
                                datas.Add(data);

                                data = new Data()
                                {
                                    cell = "M1",
                                    css = "imported_class10",
                                    format = "text",
                                    value = key.LotNumber,
                                };
                                datas.Add(data);
                                #endregion


                                int row = 2;

                                foreach (var item in group)
                                {
                                    if (string.IsNullOrWhiteSpace(item.LotNumber)) throw new SystemException("【批號】不能為空!");
                                    string[] QcItemTempletData = null;

                                    if (item.QcItemTemplet != "")
                                    {
                                        string lastStr = item.QcItemTemplet.Substring(item.QcItemTemplet.Length - 1);
                                        switch (lastStr)
                                        {
                                            case "1":
                                                QcItemTempletData = new string[] {
                                                    "B01U01z126:面型量測單:Panasonic UA3P-300(量測中心)",
                                                    "B01Q01z124:外徑(HQV)量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01A01z128:邊厚/平行度量測單:尼康高度计-1 MF-501 0-50MM 0-35MM",
                                                    "B01P01z130:光學偏芯量測單:偏芯仪 佐庆TRIOPTICS"
                                                };
                                                break;
                                            case "2":
                                                QcItemTempletData = new string[] {
                                                    "B01U01z126:面型量測單:Panasonic UA3P-300(量測中心)",
                                                    "B01Q01z124:外徑(HQV)量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01A01z128:邊厚/平行度 量測單:尼康高度计-1 MF-501 0-50MM 0-35MM",
                                                    "B01Q01z131:同軸偏芯量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01MI1z127:外徑(基恩士)量測單:基恩士測維儀 LS-9030MT"
                                                };
                                                break;
                                            case "3":
                                                QcItemTempletData = new string[] {
                                                    "B01U01z126:面型量測單:Panasonic UA3P-300(量測中心)",
                                                    "B01Q01z124:外徑(HQV)量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01A01z128:邊厚/平行度 量測單:尼康高度计-1 MF-501 0-50MM 0-35MM",
                                                    "B01P01z130:光學偏芯量測單:偏芯仪 佐庆TRIOPTICS",
                                                    "B01MI1z127:外徑(基恩士)量測單:基恩士測維儀 LS-9030MT"
                                                };
                                                break;
                                            case "4":
                                                QcItemTempletData = new string[] {
                                                    "B01U01z126:面型量測單:Panasonic UA3P-300(量測中心)",
                                                    "B01Q01z124:外徑(HQV)量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01A01z128:邊厚/平行度 量測單:尼康高度计-1 MF-501 0-50MM 0-35MM",
                                                    "B01Q01z131:同軸偏芯量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01MI1z127:外徑(基恩士)量測單:基恩士測維儀 LS-9030MT",
                                                    "B01I01z133:剪口量測單:2.5D Bao-I 2.5次元"
                                                };
                                                break;
                                            case "5":
                                                QcItemTempletData = new string[] {
                                                    "B01U01z126:面型量測單:Panasonic UA3P-300(量測中心)",
                                                    "B01Q01z124:外徑(HQV)量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01A01z128:邊厚/平行度 量測單:尼康高度计-1 MF-501 0-50MM 0-35MM",
                                                    "B01P01z130:光學偏芯量測單:偏芯仪 佐庆TRIOPTICS",
                                                    "B01MI1z127:外徑(基恩士)量測單:基恩士測維儀 LS-9030MT",
                                                    "B01I01z133:剪口量測單:2.5D Bao-I 2.5次元"
                                                };
                                                break;
                                            case "6":
                                                QcItemTempletData = new string[] {
                                                    "B01U01z126:面型量測單:Panasonic UA3P-300(量測中心)",
                                                    "B01Q01z124:外徑(HQV)量測單:QV-1 Mitutoyo H302P1L-D",
                                                    "B01A01z128:邊厚/平行度 量測單:尼康高度计-1 MF-501 0-50MM 0-35MM",
                                                    "B01P01z130:光學偏芯量測單:偏芯仪 佐庆TRIOPTICS",
                                                    "B01MI1z127:外徑(基恩士)量測單:基恩士測維儀 LS-9030MT"
                                                };
                                                break;
                                        }


                                        foreach (var item1 in QcItemTempletData)
                                        {
                                            #region //設定量測項目、備註、上下限
                                            data = new Data()
                                            {
                                                cell = "A" + row,
                                                css = "",
                                                format = "common",
                                                value = item1.Split(':')[0],
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "B" + row,
                                                css = "",
                                                format = "text",
                                                value = "",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "C" + row,
                                                css = "",
                                                format = "text",
                                                value = item1.Split(':')[1],
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "D" + row,
                                                css = "",
                                                format = "text",
                                                value = item1.Split(':')[1],
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "E" + row,
                                                css = "",
                                                format = "text",
                                                value = item1.Split(':')[2],
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "F" + row,
                                                css = "",
                                                format = "text",
                                                value = "0",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "G" + row,
                                                css = "",
                                                format = "text",
                                                value = "1",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "H" + row,
                                                css = "",
                                                format = "text",
                                                value = "1",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "I" + row,
                                                css = "",
                                                format = "text",
                                                value = "",
                                            };
                                            datas.Add(data);

                                            row++;
                                            #endregion
                                        }
                                    }
                                    else if (item.QcItemNo != "")
                                    {
                                        #region //設定量測項目、備註、上下限
                                        data = new Data()
                                        {
                                            cell = "A" + row,
                                            css = "",
                                            format = "common",
                                            value = item.QcItemNo.Split(':')[0],
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "B" + row,
                                            css = "",
                                            format = "text",
                                            value = "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "C" + row,
                                            css = "",
                                            format = "text",
                                            value = item.QcItemNo.Split(':')[1],
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "D" + row,
                                            css = "",
                                            format = "text",
                                            value = item.QcItemRemark,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "E" + row,
                                            css = "",
                                            format = "text",
                                            value = item.QcItemNo.Split(':')[2],
                                            //value = item.Machine != null ? item.Machine : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "F" + row,
                                            css = "",
                                            format = "text",
                                            value = "0",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "G" + row,
                                            css = "",
                                            format = "text",
                                            value = "1",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "H" + row,
                                            css = "",
                                            format = "text",
                                            value = "1",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "I" + row,
                                            css = "",
                                            format = "text",
                                            value = "",
                                        };
                                        datas.Add(data);

                                        row++;
                                        #endregion
                                    }
                                    else
                                    {
                                        throw new SystemException("【量測項目資料】不能為空!");
                                    }






                                }

                                #region //整合Spreadsheet格式
                                List<Sheets> sheetss = new List<Sheets>();

                                Sheets sheets = new Sheets()
                                {
                                    name = "sheet1",
                                    data = datas
                                };
                                sheetss.Add(sheets);

                                #region //更新至QcRecord SpreadsheetData
                                SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                {
                                    sheets = sheetss
                                };

                                string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion

                            }


                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                qcRecordIds = String.Join(",", qcRecordIds)
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
        #region //UpdateQcRecord -- 更新量測記錄單頭資料 -- Ann 2023-02-27
        public string UpdateQcRecord(int QcRecordId, int QcNoticeId, int QcTypeId, int MoId, int MoProcessId, string Remark, int CurrentFileId, string QcRecordFile, string InputType, string ServerPath, string ServerPath2, string SupportAqFlag, string ResolveFile, string ResolveFileJson, string QcRecordFileByNas)
        {
            try
            {
                if (InputType.Length < 0) throw new SystemException("【輸入方式】不能為空!");
                if (SupportAqFlag.Length < 0) throw new SystemException("【是否支援品異單】不能為空!");
                //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Cavity") throw new SystemException("【穴號模式】只能由全吋檢或成型工程檢使用!!");
                //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Picture") throw new SystemException("【面型圖模式】只能由全吋檢或成型工程檢使用!!");

                string ErpDbName = "";
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //檢查此量測資料是否已經有上傳數據
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcMeasureData
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcMeasureDataResult.Count() > 0) throw new SystemException("此量測資料已有上傳數據，無法更改!");
                            #endregion

                            #region //判斷量測記錄資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CheckQcMeasureData
                                    FROM MES.QcRecord
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");

                            string CheckQcMeasureData = "";
                            foreach (var item in QcRecordResult)
                            {
                                if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "S") throw new SystemException("量測單據狀態無法更改!!");
                                CheckQcMeasureData = item.CheckQcMeasureData;
                            }
                            #endregion

                            #region //判斷MES製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT (b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(varchar(10), a.WoSeq) + ')') WoErpFullNo
                                    , b.WoErpPrefix, b.WoErpNo
                                    , c.MtlItemNo, c.MtlItemName
                                    FROM MES.ManufactureOrder a 
                                    INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    WHERE MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("MES製令資料錯誤!");

                            string WoErpFullNo = "";
                            string MtlItemNo = "";
                            string MtlItemName = "";
                            string WoErpPrefix = "";
                            string WoErpNo = "";
                            foreach (var item in result)
                            {
                                WoErpFullNo = item.WoErpFullNo;
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                                WoErpPrefix = item.WoErpPrefix;
                                WoErpNo = item.WoErpNo;
                            }
                            #endregion

                            #region //判斷ERP製令狀態是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TA011, TA013
                                    FROM MOCTA
                                    WHERE TA001 = @TA001
                                    AND TA002 = @TA002";
                            dynamicParameters.Add("TA001", WoErpPrefix);
                            dynamicParameters.Add("TA002", WoErpNo);

                            var MOCTAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (MOCTAResult.Count() <= 0) throw new SystemException("ERP製令資料錯誤!!");

                            foreach (var item in MOCTAResult)
                            {
                                if (SupportAqFlag == "Y" && (item.TA011 == "Y" || item.TA011 == "y" || item.TA013 == "V")) throw new SystemException("ERP製令狀態無法開立!!");
                            }
                            #endregion

                            #region //判斷量測類型資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                    FROM QMS.QcType a 
                                    WHERE a.QcTypeId = @QcTypeId";
                            dynamicParameters.Add("QcTypeId", QcTypeId);

                            var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                            string QcTypeNo = "";
                            string QcTypeName = "";
                            foreach (var item in QcTypeResult)
                            {
                                if (item.SupportProcessFlag == "Y" && MoProcessId <= 0) throw new SystemException("【量測製程】不能為空!");
                                QcTypeNo = item.QcTypeNo;
                                QcTypeName = item.QcTypeName;
                            }
                            #endregion

                            if (MoProcessId > 0)
                            {
                                #region //判斷製程資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.MoProcess
                                        WHERE MoProcessId = @MoProcessId";
                                dynamicParameters.Add("MoProcessId", MoProcessId);

                                var MoProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                if (MoProcessResult.Count() <= 0) throw new SystemException("製令製程資料錯誤!");
                                #endregion
                            }

                            #region //建立CheckQcMeasureData
                            if (InputType == "Cavity" || InputType == "Picture")
                            {
                                if (QcRecordFile.Length > 0) CheckQcMeasureData = "Y";
                            }
                            #endregion

                            #region //UPDATE MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    QcNoticeId = @QcNoticeId,
                                    QcTypeId = @QcTypeId,
                                    InputType = @InputType,
                                    MoId = @MoId,
                                    MoProcessId = @MoProcessId,
                                    Remark = @Remark,
                                    CurrentFileId = @CurrentFileId,
                                    CheckQcMeasureData = @CheckQcMeasureData,
                                    SupportAqFlag = @SupportAqFlag,
                                    ResolveFileJson = @ResolveFileJson,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId = QcNoticeId <= 0 ? (int?)null : QcNoticeId,
                                    QcTypeId,
                                    InputType,
                                    MoId,
                                    MoProcessId = MoProcessId <= 0 ? (int?)null : MoProcessId,
                                    Remark,
                                    CurrentFileId,
                                    CheckQcMeasureData,
                                    SupportAqFlag,
                                    ResolveFileJson,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新FILE資訊

                            #region //先將原本資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.QcRecordFile
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INSERT MES.QcRecordFile
                            if (QcRecordFile.Length > 0)
                            {
                                var qcRecordFileList = QcRecordFile.Split(',');

                                foreach (var qcRecordFileId in qcRecordFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileId = qcRecordFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT MES.QcRecordFile By NAS
                            if (QcRecordFileByNas.Length > 0)
                            {
                                JObject uploadFileJson = JObject.Parse(QcRecordFileByNas);

                                foreach (var item in uploadFileJson["uploadFileInfo"])
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
                                            FileType = "other",
                                            PhysicalPath = item["FilePath"].ToString(),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT ResolveFile
                            if (ResolveFile.Length > 0)
                            {
                                var resolveFileList = ResolveFile.Split(',');

                                foreach (var resolveFileId in resolveFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = "resolve",
                                            FileId = resolveFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion
                            #endregion

                            #region //將量測EXCEL先存回實體路徑並重新讀取
                            string FileName = "";
                            if (CurrentFileId > 0)
                            {
                                #region //取得原本加密檔案資料，並重新上傳
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId, a.FileContent, a.FileName, a.FileExtension, a.FileSize, a.ClientIP, a.Source
                                        FROM BAS.[File] a
                                        WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", CurrentFileId);

                                var FileResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in FileResult)
                                {
                                    FileName = item.FileName;
                                    ServerPath = Path.Combine(ServerPath, item.FileName + "-" + item.FileId + item.FileExtension);
                                    byte[] fileContent = (byte[])item.FileContent;
                                    File.WriteAllBytes(ServerPath, fileContent); // Requires System.IO

                                    #region //將檔案重新上傳
                                    System.Net.WebClient wc = new System.Net.WebClient();
                                    byte[] newFilebytes = wc.DownloadData(ServerPath);

                                    dynamicParameters = new DynamicParameters();
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
                                            CompanyId = CurrentCompany,
                                            item.FileName,
                                            FileContent = newFilebytes,
                                            item.FileExtension,
                                            item.FileSize,
                                            item.ClientIP,
                                            item.Source,
                                            DeleteStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    int newFileId = -1;
                                    foreach (var item2 in insertResult)
                                    {
                                        newFileId = item2.FileId;
                                    }
                                    #endregion

                                    #region //UPDATE原本FileId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcRecord SET
                                            CurrentFileId = @CurrentFileId,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CurrentFileId = newFileId,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QcRecordId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除實體檔案
                                    File.Delete(ServerPath);
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            if ((CheckQcMeasureData == "Y" || CheckQcMeasureData == "P") && CurrentFileId > 0 && (QcTypeNo == "PVTQC" || QcTypeNo == "TQC"))
                            {
                                SendQcMail(sqlConnection, QcRecordId, WoErpFullNo, MtlItemNo, MtlItemName, QcTypeNo, QcTypeName
                                    , Remark, FileName, ServerPath2, MoId);
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

        #region //VoidQcRecord -- 作廢量測記錄單頭資料 -- Ann 2023-02-27
        public string VoidQcRecord(int QcRecordId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測記錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CheckQcMeasureData
                                FROM MES.QcRecord
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("無法作廢，查無量測記錄資料!");
                                          
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                CheckQcMeasureData = @CheckQcMeasureData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CheckQcMeasureData = "V",
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //Update MoProcessQty
        public void UpdateMoQtyInformation(int MoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //更新MoProcess PassQty,NgQty,ScrapQty,InputQty
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcess
                                   SET TotalPassQty = y.PassQty,
                                       TotalNgQty = y.NgQty,
	                                   TotalScrapQty = y.ScrapQty,
	                                   TotalInputQty = y.InputQty
                                  FROM MES.MoProcess x,
	                                   (SELECT t.MoProcessId,
			                                   t.ProcessAlias,
			                                   SUM(t.PassQty) PassQty,
			                                   SUM(t.NgQty) NgQty,
			                                   SUM(t.ScrapQty) ScrapQty,
			                                   SUM(t.PassQty) + SUM(t.NgQty) + SUM(t.ScrapQty) InputQty
		                                  FROM (
				                                SELECT a.MoProcessId,a.ProcessAlias,
					                                   CASE WHEN b.ProdStatus = 'P' THEN SUM(b.StationQty) ELSE 0 END PassQty,
					                                   CASE WHEN b.ProdStatus = 'F' THEN SUM(b.StationQty) ELSE 0 END NgQty,
					                                   CASE WHEN b.ProdStatus = 'S' THEN SUM(b.StationQty) ELSE 0 END ScrapQty
				                                  FROM MES.MoProcess a
					                                   INNER JOIN MES.BarcodeProcess b ON a.MoProcessId = b.MoProcessId
				                                 GROUP BY a.MoProcessId,a.ProcessAlias,b.ProdStatus) t
		                                 GROUP BY t.MoProcessId,t.ProcessAlias) y
                                 WHERE x.MoProcessId  = y.MoProcessId
                                   AND x.MoId = @MoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MoId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        #endregion

                        #region //找出所有MoProcess依順序更新WipQty
                        //目前無法有邏輯推算每站在製數量, 等確定邏輯再補
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcess
                                   SET WipQty = ISNULL(b.WipQty,0)
                                  FROM MES.MoProcess a
                                       LEFT JOIN (SELECT x.CurrentMoProcessId,SUM(x.BarcodeQty) WipQty
	                                                FROM MES.Barcode x
				                                   WHERE x.BarcodeStatus = 1
				                                     AND EXISTS(SELECT 1
					                                              FROM MES.BarcodeProcess z
								                                 WHERE x.BarcodeId = z.BarcodeId
								                                   AND z.MoId = x.MoId)
				                                   GROUP BY x.CurrentMoProcessId) b ON a.MoProcessId = b.CurrentMoProcessId
                                 WHERE a.MoId = @MoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MoId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新Mo Scrap Quantity
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ManufactureOrder
                                   SET ScrapQty = y.ScrapQty
                                  FROM MES.ManufactureOrder x,
                                       (SELECT a.MoId,ISNULL(SUM(b.TotalScrapQty),0) ScrapQty
                                          FROM MES.ManufactureOrder a
                                               INNER JOIN MES.MoProcess b ON a.MoId = b.MoId
	                                     GROUP BY a.MoId) y
                                 WHERE x.MoId = y.MoId
                                   AND x.MoId = @MoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MoId
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
        }
        #endregion 

        #region //UpdateResolveFileJson -- 暫存量測檔案解析JSON -- Ann 2023-11-14
        public string UpdateResolveFileJson(int QcRecordId, string ResolveFileJson)
        {
            try
            {
                if (ResolveFileJson.Length < 0) throw new SystemException("【量測解析JSON】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //檢查此量測資料是否已經有上傳數據
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcMeasureDataResult.Count() > 0) throw new SystemException("此量測資料已有上傳數據，無法更改!");
                        #endregion

                        #region //判斷量測記錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CheckQcMeasureData, MoId
                                FROM MES.QcRecord
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");

                        string CheckQcMeasureData = "";
                        int MoId = -1;
                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "S") throw new SystemException("量測單據狀態無法更改!!");
                            CheckQcMeasureData = item.CheckQcMeasureData;
                            MoId = item.MoId;
                        }
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                ResolveFileJson = @ResolveFileJson,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ResolveFileJson,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //UpdateCheckGoods -- 確認收貨狀態 -- Ann 2024-1-24
        public string UpdateCheckGoods(int QcRecordId, string CheckQcMeasureData, string Disallowance)
        {
            try
            {
                if (CheckQcMeasureData.Length < 0) throw new SystemException("【更改狀態】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo
                                FROM BAS.Company a 
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //取得收件人員資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo, a.UserName
                                FROM BAS.[User] a 
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var CheckUserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CheckUserResult.Count() <= 0) throw new SystemException("收件人員資料錯誤!!");

                        string UserNo = "";
                        string UserName = "";
                        foreach (var item in CheckUserResult)
                        {
                            UserNo = item.UserNo;
                            UserName = item.UserName;
                        }
                        #endregion

                        #region //判斷量測記錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId, a.CheckQcMeasureData, ISNULL(FORMAT(a.ReceiptDate, 'yyyy-MM-dd'), '') ReceiptDate
                                , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') DocCreateDate
                                , b.WoSeq, b.ModeId
                                , c.WoErpPrefix, c.WoErpNo
                                , d.ProcessAlias
                                , e.DepartmentId
                                , f.DepartmentNo, f.DepartmentName
                                , g.QcTypeNo, g.QcTypeName
                                , h.MtlItemNo, h.MtlItemName
                                , i.ModeName
                                FROM MES.QcRecord a 
                                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                LEFT JOIN MES.MoProcess d ON a.MoProcessId = d.MoProcessId
                                INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId
                                INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                                INNER JOIN QMS.QcType g ON a.QcTypeId = g.QcTypeId
                                INNER JOIN PDM.MtlItem h ON c.MtlItemId = h.MtlItemId
                                INNER JOIN MES.ProdMode i ON b.ModeId = i.ModeId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");

                        string WoErpPrefix = "";
                        string WoErpNo = "";
                        string WoSeq = "";
                        string QcTypeNo = "";
                        string QcTypeName = "";
                        string ProcessAlias = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ReceiptDate = "";
                        string MtlItemNo = "";
                        string MtlItemName = "";
                        int DepartmentId = -1;
                        int ModeId = -1;
                        string ModeName = "";
                        string DocCreateDate = "";
                        foreach (var item in QcRecordResult)
                        {
                            if (CheckQcMeasureData == "C" && item.CheckQcMeasureData != "A") throw new SystemException("單據非送測未確認狀態!!");
                            if (CheckQcMeasureData == "F" && item.CheckQcMeasureData != "A") throw new SystemException("單據非送測未確認狀態!!");
                            if (CheckQcMeasureData == "S" && (item.CheckQcMeasureData != "A" && item.CheckQcMeasureData != "C")) throw new SystemException("單據非送測未確認狀態!!");

                            WoErpPrefix = item.WoErpPrefix;
                            WoErpNo = item.WoErpNo;
                            WoSeq = item.WoSeq.ToString();
                            QcTypeNo = item.QcTypeNo;
                            QcTypeName = item.QcTypeName;
                            ProcessAlias = item.ProcessAlias;
                            DepartmentNo = item.DepartmentNo;
                            DepartmentName = item.DepartmentName;
                            ReceiptDate = item.ReceiptDate.ToString();
                            MtlItemNo = item.MtlItemNo;
                            MtlItemName = item.MtlItemName;
                            DepartmentId = item.DepartmentId;
                            ModeId = item.ModeId;
                            ModeName = item.ModeName;
                            DocCreateDate = item.DocCreateDate.ToString();
                        }
                        #endregion

                        #region //確認收貨及開始量測欄位
                        string QcStartDate = "";
                        string CurrentDatetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        switch (CheckQcMeasureData)
                        {
                            case "C":
                                ReceiptDate = CurrentDatetime;
                                break;
                            case "S":
                                QcStartDate = CurrentDatetime;
                                if (ReceiptDate == "")
                                {
                                    ReceiptDate = QcStartDate;
                                }
                                break;
                            case "F":
                                ReceiptDate = CurrentDatetime;
                                break;
                            default:
                                throw new SystemException("單據狀態錯誤!!");
                        }
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                CheckQcMeasureData = @CheckQcMeasureData,
                                ReceiptDate = @ReceiptDate,
                                QcStartDate = @QcStartDate,
                                DisallowanceReason = @Disallowance,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CheckQcMeasureData,
                                ReceiptDate,
                                QcStartDate,
                                Disallowance,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //MAMO推播通知
                        string Content = "";

                        #region //取得送測條碼清單
                        string barcodeInfoDesc = "| 條碼 | 刻字 |\n| :-- | :-- |\n";
                        int count = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.BarcodeId
                                , ISNULL(c.ItemValue, '') ItemValue
                                , d.BarcodeNo
                                FROM MES.QcReceiptBarcode a 
                                INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                LEFT JOIN MES.BarcodeAttribute c ON a.BarcodeId = c.BarcodeId AND c.MoId = b.MoId AND c.ItemNo = 'Lettering'
                                INNER JOIN MES.Barcode d ON a.BarcodeId = d.BarcodeId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcReceiptBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in QcReceiptBarcodeResult)
                        {
                            barcodeInfoDesc += "| " + item.BarcodeNo + " | " + item.ItemValue + " |\n";
                            count++;
                        }
                        #endregion

                        if (CheckQcMeasureData == "C")
                        {
                            Content = "### 【量測確認收件通知】\n" +
                                            "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + DocCreateDate + "\n" +
                                            "- 收件人員: " + UserNo + " " + UserName + "\n" +
                                            "- 確認收件時間: " + CreateDate.ToString() + "\n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        }
                        else if (CheckQcMeasureData == "F")
                        {
                            Content = "### 【量測單據駁回通知】\n" +
                                            "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + DocCreateDate + "\n" +
                                            "- 駁回人員: " + UserNo + " " + UserName + "\n" +
                                            "- 駁回時間: " + CreateDate.ToString() + "\n" +
                                            "- 駁回原因: " + Disallowance + "\n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        }
                        else if (CheckQcMeasureData == "S")
                        {
                            Content = "### 【開始量測通知】\n" +
                                            "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + DocCreateDate + "\n" +
                                            "- 量測人員: " + UserNo + " " + UserName + "\n" +
                                            "- 開始量測時間: " + CreateDate.ToString() + "\n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        }
                        else
                        {
                            throw new SystemException("確認收貨類型錯誤!!");
                        }

                        #region //確認推播群組
                        string SendId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                                FROM MES.QcMamoChannel a
                                WHERE a.ModeId = @ModeId
                                AND PushType = @PushType";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("PushType", CheckQcMeasureData);

                        var QcMamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMamoChannelResult.Count() <= 0) throw new SystemException("收貨確認/駁回推播群組資料錯誤!!<br>請確認生產模式【" + ModeName + "】是否已設定推播群組!!");

                        int ChannelId = -1;
                        foreach (var item in QcMamoChannelResult)
                        {
                            ChannelId = item.ChannelId;
                            SendId = ChannelId.ToString();
                        }
                        #endregion

                        #region //取得標記USER資料(原送測人員部門)
                        List<string> Tags = new List<string>();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId
                                , b.UserNo, b.UserName
                                FROM MAMO.ChannelMembers a 
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.ChannelId = @SendId
                                AND b.DepartmentId = @DepartmentId";
                        dynamicParameters.Add("SendId", SendId);
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in UserResult)
                        {
                            Tags.Add(item.UserNo);
                        }
                        #endregion

                        List<int> Files = new List<int>();

                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", SendId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
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

        #region //UpdateUrgent -- 更新急件狀態 -- Ann 2024-04-11
        public string UpdateUrgent(int QcRecordId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷確認量測單據資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CheckQcMeasureData, a.UrgentFlag
                                FROM MES.QcRecord a 
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測單據資料錯誤!");

                        string UrgentFlag = "";
                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "A" && item.CheckQcMeasureData != "C" && item.CheckQcMeasureData != "P" && item.CheckQcMeasureData != "S") throw new SystemException("單據狀態已完成量測，無法改為急件!!");
                            UrgentFlag = item.UrgentFlag;
                        }
                        #endregion

                        switch (UrgentFlag)
                        {
                            case "Y":
                                UrgentFlag = "N";
                                break;
                            case "N":
                                UrgentFlag = "Y";
                                break;
                            default:
                                throw new SystemException("量測急件狀態有誤!!");
                        }

                        #region //UPDATE MES.QcRecord UrgentFlag
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                UrgentFlag = @UrgentFlag,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UrgentFlag,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //UpdateQcGoodsReceipt -- 更新進貨檢量測單據 -- Ann 2024-04-25
        public string UpdateQcGoodsReceipt(int QcRecordId, int GrDetailId, string InputType, string QcRecordFile, string ResolveFile, string ResolveFileJson, string Remark, string ServerPath, string ServerPath2)
        {
            try
            {
                string ErpDbName = "";
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (InputType.Length < 0) throw new SystemException("【輸入方式】不能為空!");

                            #region //確認進貨單身資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.CloseStatus
                                    , b.ConfirmStatus
                                    FROM SCM.GrDetail a 
                                    INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");

                            foreach (var item in GrDetailResult)
                            {
                                if (item.CloseStatus != "N") throw new SystemException("此進貨單身已結案!!");
                                if (item.ConfirmStatus != "N") throw new SystemException("此進貨單身已確認!!");
                            }
                            #endregion

                            #region //UPDATE MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    InputType = @InputType,
                                    Remark = @Remark,
                                    ResolveFileJson = @ResolveFileJson,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    InputType,
                                    Remark,
                                    ResolveFileJson,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //UPDATE MES.QcGoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcGoodsReceipt SET
                                    GrDetailId = @GrDetailId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    GrDetailId,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新FILE資訊

                            #region //先將原本資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.QcRecordFile
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            if (QcRecordFile.Length > 0)
                            {
                                var qcRecordFileList = QcRecordFile.Split(',');

                                foreach (var qcRecordFileId in qcRecordFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileId = qcRecordFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }

                            #region //INSERT ResolveFile
                            if (ResolveFile.Length > 0)
                            {
                                var resolveFileList = ResolveFile.Split(',');

                                foreach (var resolveFileId in resolveFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = "resolve",
                                            FileId = resolveFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion
                            #endregion

                            #region //統整回傳前端資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId, a.CurrentFileId
                                    , (
                                        SELECT x.FileId
                                        FROM MES.QcRecordFile x
                                        WHERE x.QcRecordId = a.QcRecordId
                                        FOR JSON PATH, ROOT('data')
                                    ) QcRecordFile
                                    FROM MES.QcRecord a
                                    WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                data = QcRecordResult
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

        #region //UpdateQcMachineModePlanningSort -- 更新量測機台排程順序 -- Ann 2024-08-22
        public string UpdateQcMachineModePlanningSort(string PlanningSort)
        {
            try
            {
                if (PlanningSort.Length < 0) throw new SystemException("【排序資料】不能為空!");

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string[] SortArray = PlanningSort.Split(',');
                        foreach (var qmmpId in SortArray)
                        {
                            int currentSort = Array.IndexOf(SortArray, qmmpId) + 1;

                            #region //確認量測機台排程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.Sort, a.QmmDetailId, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                                    , b.QrpId, b.EstimatedMeasurementTime
                                    , c.QcRecordId, c.CheckQcMeasureData
                                    FROM QMS.QcMachineModePlanning a 
                                    INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                    INNER JOIN MES.QcRecord c ON b.QcRecordId = c.QcRecordId
                                    WHERE a.QmmpId = @QmmpId";
                            dynamicParameters.Add("QmmpId", qmmpId);

                            var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcMachineModePlanningResult.Count() <= 0) throw new SystemException("量測機台排程資料錯誤!!");

                            int OriSort = 0;
                            string CheckQcMeasureData = "";
                            int QcRecordId = -1;
                            int QrpId = -1;
                            double EstimatedMeasurementTime = 0;
                            int QmmDetailId = -1;
                            foreach (var item in QcMachineModePlanningResult)
                            {
                                if (item.CreateDate != DateTime.Now.ToString("yyyy-MM-dd"))
                                {
                                    throw new SystemException("不能夠異動非當日之排程資料!!");
                                }

                                OriSort = item.Sort;
                                CheckQcMeasureData = item.CheckQcMeasureData;
                                QcRecordId = item.QcRecordId;
                                QrpId = item.QrpId;
                                EstimatedMeasurementTime = item.EstimatedMeasurementTime;
                                QmmDetailId = item.QmmDetailId;
                            }
                            #endregion

                            if (OriSort != currentSort)
                            {
                                #region //確認量測單據狀態
                                if (CheckQcMeasureData != "N" && CheckQcMeasureData != "A" && CheckQcMeasureData != "C")
                                {
                                    throw new SystemException("影響量測單號【" + QcRecordId + "】原順序【" + OriSort + "】狀態錯誤，無法更改排程!!");
                                }
                                #endregion

                                #region //取得新排程初始開始日期
                                DateTime firstStartDate = DateTime.Now;
                                if (currentSort != 1)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.EstimatedEndDate
                                            FROM QMS.QcMachineModePlanning a 
                                            WHERE a.Sort = @Sort
                                            AND a.QmmDetailId = @QmmDetailId
                                            AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow";
                                    dynamicParameters.Add("Sort", currentSort - 1);
                                    dynamicParameters.Add("QmmDetailId", QmmDetailId);
                                    dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                                    var EstimatedEndDateResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (EstimatedEndDateResult.Count() <= 0) throw new SystemException("取得初始開始日期時錯誤!!");

                                    foreach (var item in EstimatedEndDateResult)
                                    {
                                        firstStartDate = item.EstimatedEndDate;
                                    }
                                }
                                #endregion

                                #region //確認此量測單據是否已經有其他機型已排定排程，若有則需檢核時間順序
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.QmmDetailId, a.EstimatedEndDate
                                        , d.MachineDesc
                                        FROM QMS.QcMachineModePlanning a 
                                        INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                        INNER JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                        INNER JOIN MES.Machine d ON c.MachineId = d.MachineId
                                        WHERE b.QcRecordId = @QcRecordId
                                        AND a.QrpId != @QrpId
                                        AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow
                                        ORDER BY a.EstimatedEndDate DESC";
                                dynamicParameters.Add("QcRecordId", QcRecordId);
                                dynamicParameters.Add("QrpId", QrpId);
                                dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                                var QcMachineModePlanningResult2 = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in QcMachineModePlanningResult2)
                                {
                                    DateTime finalEndDate = item.EstimatedEndDate;
                                    if (firstStartDate < finalEndDate)
                                    {
                                        firstStartDate = finalEndDate;
                                        //throw new SystemException("此次量測單號【" + QcRecordId + "】排程開始時間【" + firstStartDate + "】不可早於上個排程【" + item.MachineDesc + "】結束時間【" + finalEndDate.ToString("yyyy-MM-dd HH:mm:ss") + "】!!");
                                    }
                                }
                                #endregion

                                #region //取得新排程之後的所有異動排程資料
                                List<int> OriQmmpIdSortArray = new List<int>();
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QmmpId, a.Sort, a.EstimatedStartDate, a.EstimatedEndDate
                                        , b.EstimatedMeasurementTime
                                        , c.QcStartDate, c.QcRecordId
                                        FROM QMS.QcMachineModePlanning a 
                                        INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                        INNER JOIN MES.QcRecord c ON b.QcRecordId = c.QcRecordId
                                        WHERE a.Sort >= @Sort
                                        AND a.QmmDetailId = @QmmDetailId
                                        AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow
                                        ORDER BY a.Sort";
                                dynamicParameters.Add("Sort", currentSort);
                                dynamicParameters.Add("QmmDetailId", QmmDetailId);
                                dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                                var PlanningSortCheckResult2 = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in PlanningSortCheckResult2)
                                {
                                    if (item.QcStartDate != null)
                                    {
                                        throw new SystemException("【原排序: " + item.Sort + "】已開始量測，無法變更其排程順序!!");
                                    }

                                    OriQmmpIdSortArray.Add(item.QmmpId);
                                }
                                #endregion

                                #region //Remove原先排程順序
                                if (OriQmmpIdSortArray.Count() > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                            Sort = NULL,
                                            EstimatedStartDate = NULL,
                                            EstimatedEndDate = NULL
                                            WHERE QmmpId IN (" + string.Join(",", OriQmmpIdSortArray) + ")";

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                }
                                #endregion

                                #region //先排序異動排程
                                DateTime EstimatedEndDate = firstStartDate.AddMinutes(EstimatedMeasurementTime);
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                        Sort = @Sort,
                                        EstimatedStartDate = @EstimatedStartDate,
                                        EstimatedEndDate = @EstimatedEndDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QmmpId = @QmmpId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        Sort = currentSort,
                                        EstimatedStartDate = firstStartDate,
                                        EstimatedEndDate,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QmmpId = qmmpId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //Reset原先排程順序
                                int i = 1;
                                DateTime currentEstimatedEndDate = EstimatedEndDate;
                                foreach (var oriQmmpId in OriQmmpIdSortArray)
                                {
                                    if (oriQmmpId == Convert.ToInt32(qmmpId)) continue;

                                    #region //計算新的開始日期
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT b.EstimatedMeasurementTime
                                            FROM QMS.QcMachineModePlanning a 
                                            INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                            WHERE a.QmmpId = @QmmpId";
                                    dynamicParameters.Add("QmmpId", oriQmmpId);

                                    var GetEstimatedMeasurementTimeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (GetEstimatedMeasurementTimeResult.Count() <= 0) throw new SystemException("【量測機台排程ID: " + qmmpId + "】查無此排程資料!!");

                                    double currentEstimatedMeasurementTime = 0;
                                    foreach (var item in GetEstimatedMeasurementTimeResult)
                                    {
                                        currentEstimatedMeasurementTime = item.EstimatedMeasurementTime;
                                    }
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                                Sort = @Sort,
                                                EstimatedStartDate = @EstimatedStartDate,
                                                EstimatedEndDate = @EstimatedEndDate,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE QmmpId = @QmmpId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            Sort = currentSort + i,
                                            EstimatedStartDate = currentEstimatedEndDate,
                                            EstimatedEndDate = currentEstimatedEndDate.AddMinutes(currentEstimatedMeasurementTime),
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QmmpId = oriQmmpId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                    i++;

                                    currentEstimatedEndDate = currentEstimatedEndDate.AddMinutes(currentEstimatedMeasurementTime);
                                }
                                #endregion

                                break;
                            }
                        }

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

        #region //UpdateEstimatedMeasurementTime -- 更新量測單據排程預計量測時間 -- Ann 2024-08-23
        public string UpdateEstimatedMeasurementTime(string UploadJson)
        {
            try
            {
                if (UploadJson.Length <= 0) throw new SystemException("【排程資料】不能為空!!");

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        JObject uploadJson = JObject.Parse(UploadJson);

                        foreach (var item in uploadJson["uploadInfo"])
                        {
                            string QrpId = item["QrpId"].ToString();
                            double EstimatedMeasurementTime = Convert.ToDouble(item["EstimatedMeasurementTime"]);

                            #region //確認量測單據排程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId 
                                    , b.CheckQcMeasureData
                                    , c.QcMachineModeName
                                    FROM QMS.QcRecordPlanning a 
                                    INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                    INNER JOIN QMS.QcMachineMode c ON a.QcMachineModeId = c.QcMachineModeId
                                    WHERE a.QrpId = @QrpId";
                            dynamicParameters.Add("QrpId", QrpId);

                            var QcRecordPlanningResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcRecordPlanningResult.Count() <= 0) throw new SystemException("量測單據排程資料錯誤!!");

                            int QcRecordId = -1;
                            foreach (var item2 in QcRecordPlanningResult)
                            {
                                if (item2.CheckQcMeasureData != "N" && item2.CheckQcMeasureData != "A" && item2.CheckQcMeasureData != "C")
                                {
                                    throw new SystemException("量測單號【" + item2.QcRecordId + "】狀態無法更改排程資料!!");
                                }

                                if (EstimatedMeasurementTime <= 0)
                                {
                                    throw new SystemException("機型【" + item2.QcMachineModeName + "】預估量測時間不可為0或小於0!!");
                                }

                                QcRecordId = item2.QcRecordId;
                            }
                            #endregion

                            #region //確認尚未排定排程
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcMachineModePlanning a 
                                    WHERE a.QrpId = @QrpId";
                            dynamicParameters.Add("QrpId", QrpId);

                            var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcMachineModePlanningResult.Count() > 0) throw new SystemException("此量測機型已排定排程，請先將排程資料刪除後再進行預計量測時間修改!!");
                            #endregion

                            #region //更改預計量測時間
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.QcRecordPlanning SET 
                                    EstimatedMeasurementTime = @EstimatedMeasurementTime,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QrpId = @QrpId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    EstimatedMeasurementTime,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QrpId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

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

        #region //UpdatePlanningSpreadsheetDate -- 更新重置排程暫存Spreadsheet -- Ann 2024-08-27
        public string UpdatePlanningSpreadsheetDate(int QcRecordId)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認量測單據資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CheckQcMeasureData
                                FROM MES.QcRecord a 
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecorIdResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcRecorIdResult.Count() <= 0) throw new SystemException("量測單據資料錯誤!!");

                        foreach (var item in QcRecorIdResult)
                        {
                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "A" && item.CheckQcMeasureData != "C")
                            {
                                throw new SystemException("量測單號【" + QcRecordId + "】狀態無法更改排程資料!!");
                            }
                        }
                        #endregion

                        #region //確認尚未排定排程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineModePlanning a 
                                INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                WHERE b.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMachineModePlanningResult.Count() > 0) throw new SystemException("此量測機型已排定排程，請先將排程資料刪除後再進行預計量測時間修改!!");
                        #endregion

                        #region //更改預計量測時間
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET 
                                PlanningSpreadsheetDate = NULL,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //UpdateCheckGoodsByPlanning -- 確認收貨狀態 -- Ann 2024-1-24
        public string UpdateCheckGoodsByPlanning(int QcRecordId, string CheckQcMeasureData)
        {
            try
            {
                if (CheckQcMeasureData.Length < 0) throw new SystemException("【更改狀態】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo
                                FROM BAS.Company a 
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //取得收件人員資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo, a.UserName
                                FROM BAS.[User] a 
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var CheckUserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CheckUserResult.Count() <= 0) throw new SystemException("收件人員資料錯誤!!");

                        string UserNo = "";
                        string UserName = "";
                        foreach (var item in CheckUserResult)
                        {
                            UserNo = item.UserNo;
                            UserName = item.UserName;
                        }
                        #endregion

                        #region //判斷量測記錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId, a.CheckQcMeasureData, ISNULL(FORMAT(a.ReceiptDate, 'yyyy-MM-dd'), '') ReceiptDate
                                , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') DocCreateDate
                                , b.WoSeq, b.ModeId
                                , c.WoErpPrefix, c.WoErpNo
                                , d.ProcessAlias
                                , e.DepartmentId
                                , f.DepartmentNo, f.DepartmentName
                                , g.QcTypeNo, g.QcTypeName
                                , h.MtlItemNo, h.MtlItemName
                                , i.ModeName
                                FROM MES.QcRecord a 
                                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                LEFT JOIN MES.MoProcess d ON a.MoProcessId = d.MoProcessId
                                INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId
                                INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                                INNER JOIN QMS.QcType g ON a.QcTypeId = g.QcTypeId
                                INNER JOIN PDM.MtlItem h ON c.MtlItemId = h.MtlItemId
                                INNER JOIN MES.ProdMode i ON b.ModeId = i.ModeId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");

                        string WoErpPrefix = "";
                        string WoErpNo = "";
                        string WoSeq = "";
                        string QcTypeNo = "";
                        string QcTypeName = "";
                        string ProcessAlias = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ReceiptDate = "";
                        string MtlItemNo = "";
                        string MtlItemName = "";
                        int DepartmentId = -1;
                        int ModeId = -1;
                        string ModeName = "";
                        string DocCreateDate = "";
                        foreach (var item in QcRecordResult)
                        {
                            WoErpPrefix = item.WoErpPrefix;
                            WoErpNo = item.WoErpNo;
                            WoSeq = item.WoSeq.ToString();
                            QcTypeNo = item.QcTypeNo;
                            QcTypeName = item.QcTypeName;
                            ProcessAlias = item.ProcessAlias;
                            DepartmentNo = item.DepartmentNo;
                            DepartmentName = item.DepartmentName;
                            ReceiptDate = item.ReceiptDate.ToString();
                            MtlItemNo = item.MtlItemNo;
                            MtlItemName = item.MtlItemName;
                            DepartmentId = item.DepartmentId;
                            ModeId = item.ModeId;
                            ModeName = item.ModeName;
                            DocCreateDate = item.DocCreateDate.ToString();
                        }
                        #endregion

                        #region //確認收貨及開始量測欄位
                        string QcStartDate = "";
                        string CurrentDatetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        switch (CheckQcMeasureData)
                        {
                            case "C":
                                ReceiptDate = CurrentDatetime;
                                break;
                            case "S":
                                QcStartDate = CurrentDatetime;
                                if (ReceiptDate == "")
                                {
                                    ReceiptDate = QcStartDate;
                                }
                                break;
                            case "F":
                                ReceiptDate = CurrentDatetime;
                                break;
                            default:
                                throw new SystemException("單據狀態錯誤!!");
                        }
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                CheckQcMeasureData = @CheckQcMeasureData,
                                ReceiptDate = @ReceiptDate,
                                QcStartDate = @QcStartDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CheckQcMeasureData,
                                ReceiptDate,
                                QcStartDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //MAMO推播通知
                        string Content = "";

                        #region //取得送測條碼清單
                        string barcodeInfoDesc = "| 條碼 | 刻字 |\n| :-- | :-- |\n";
                        int count = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.BarcodeId
                                , ISNULL(c.ItemValue, '') ItemValue
                                , d.BarcodeNo
                                FROM MES.QcReceiptBarcode a 
                                INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                LEFT JOIN MES.BarcodeAttribute c ON a.BarcodeId = c.BarcodeId AND c.MoId = b.MoId AND c.ItemNo = 'Lettering'
                                INNER JOIN MES.Barcode d ON a.BarcodeId = d.BarcodeId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcReceiptBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in QcReceiptBarcodeResult)
                        {
                            barcodeInfoDesc += "| " + item.BarcodeNo + " | " + item.ItemValue + " |\n";
                            count++;
                        }
                        #endregion

                        if (CheckQcMeasureData == "C")
                        {
                            Content = "### 【量測確認收件通知】\n" +
                                            "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + DocCreateDate + "\n" +
                                            "- 收件人員: " + UserNo + " " + UserName + "\n" +
                                            "- 確認收件時間: " + CreateDate.ToString() + "\n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        }
                        else if (CheckQcMeasureData == "F")
                        {
                            Content = "### 【量測單據駁回通知】\n" +
                                            "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + DocCreateDate + "\n" +
                                            "- 駁回人員: " + UserNo + " " + UserName + "\n" +
                                            "- 駁回時間: " + CreateDate.ToString() + "\n" +
                                            "- 駁回原因: \n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        }
                        else if (CheckQcMeasureData == "S")
                        {
                            Content = "### 【開始量測通知】\n" +
                                            "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + DocCreateDate + "\n" +
                                            "- 量測人員: " + UserNo + " " + UserName + "\n" +
                                            "- 開始量測時間: " + CreateDate.ToString() + "\n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        }
                        else
                        {
                            throw new SystemException("確認收貨類型錯誤!!");
                        }

                        #region //確認推播群組
                        string SendId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                                FROM MES.QcMamoChannel a
                                WHERE a.ModeId = @ModeId
                                AND PushType = @PushType";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("PushType", CheckQcMeasureData);

                        var QcMamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMamoChannelResult.Count() <= 0) throw new SystemException("收貨確認/駁回推播群組資料錯誤!!<br>請確認生產模式【" + ModeName + "】是否已設定推播群組!!");

                        int ChannelId = -1;
                        foreach (var item in QcMamoChannelResult)
                        {
                            ChannelId = item.ChannelId;
                            SendId = ChannelId.ToString();
                        }
                        #endregion

                        #region //取得標記USER資料(原送測人員部門)
                        List<string> Tags = new List<string>();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId
                                , b.UserNo, b.UserName
                                FROM MAMO.ChannelMembers a 
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.ChannelId = @SendId
                                AND b.DepartmentId = @DepartmentId";
                        dynamicParameters.Add("SendId", SendId);
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in UserResult)
                        {
                            Tags.Add(item.UserNo);
                        }
                        #endregion

                        List<int> Files = new List<int>();

                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", SendId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
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

        #region //UpdatePlanningConfirmStatus -- 更新計畫排程確認狀態 -- Ann 2024-09-05
        public string UpdatePlanningConfirmStatus(int QrpId, string ConfirmStatus)
        {
            try
            {
                if (ConfirmStatus != "Y" && ConfirmStatus != "N") throw new SystemException("【確認狀態】格式錯誤!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認量測單據排程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ConfirmStatus 
                                FROM QMS.QcRecordPlanning a 
                                WHERE a.QrpId = @QrpId";
                        dynamicParameters.Add("QrpId", QrpId);

                        var QcRecordPlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcRecordPlanningResult.Count() <= 0) throw new SystemException("【量測單據排程】資料錯誤!!");

                        foreach (var item in QcRecordPlanningResult)
                        {
                            if (item.ConfirmStatus == ConfirmStatus)
                            {
                                throw new SystemException("目前排程狀態錯誤!!");
                            }
                        }
                        #endregion

                        #region //確認是否至少有一筆機台排程資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineModePlanning a 
                                WHERE a.QrpId = @QrpId";
                        dynamicParameters.Add("QrpId", QrpId);

                        var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMachineModePlanningResult.Count() <= 0) throw new SystemException("需至少維護一筆排程資料!!");
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcRecordPlanning SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QrpId = @QrpId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                QrpId
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

        #region //UpdateQcGoodsReceiptLog -- 更新進貨檢驗作業資料 -- Ann 2025-02-07
        public string UpdateQcGoodsReceiptLog(int LogId, int QcGoodsReceiptId, int ReceiptQty, int AcceptQty, int ReturnQty
                    , string AcceptanceDate, string QcStatus, string QuickStatus, string Remark, string QcGoodsReceiptLogFile)
        {
            try
            {
                int rowsAffected = 0;
                ErpHelper erpHelper = new ErpHelper();

                using (TransactionScope transactionScope = new TransactionScope())
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
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        string companyNo = "";
                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {

                            #region //確認檢驗單作業資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcGoodsReceiptId
                                    FROM MES.QcGoodsReceiptLog a 
                                    WHERE a.LogId = @LogId";
                            dynamicParameters.Add("LogId", LogId);

                            var QcGoodsReceiptLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptLogResult.Count() <= 0) throw new SystemException("進貨檢驗單據錯誤!!");

                            foreach (var item in QcGoodsReceiptLogResult)
                            {
                                if (item.QcGoodsReceiptId != QcGoodsReceiptId)
                                {
                                    throw new SystemException("檢驗單ID錯誤!!");
                                }
                            }
                            #endregion

                            #region //確認進貨檢驗單據資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrDetailId 
                                    , b.CheckQcMeasureData
                                    , c.ConfirmStatus DetailConfirmStatus, c.ReceiptQty, c.AcceptQty, c.ReturnQty, c.OrigUnitPrice
                                    , d.GrId, d.ConfirmStatus
                                    FROM MES.QcGoodsReceipt a 
                                    INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                    INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                                    INNER JOIN SCM.GoodsReceipt d ON c.GrId = d.GrId
                                    WHERE a.QcGoodsReceiptId = @QcGoodsReceiptId";
                            dynamicParameters.Add("QcGoodsReceiptId", QcGoodsReceiptId);

                            var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptResult.Count() <= 0) throw new SystemException("進貨檢驗單據錯誤!!");

                            int GrDetailId = -1;
                            int GrId = -1;
                            double OrigUnitPrice = -1;
                            foreach (var item in QcGoodsReceiptResult)
                            {
                                if (QuickStatus == "N" && item.CheckQcMeasureData != "Y" && item.CheckQcMeasureData != "P")
                                {
                                    throw new SystemException("單據狀態錯誤，需先上傳量測數據後才能進行檢驗作業!!");
                                }

                                if (item.DetailConfirmStatus != "N" || item.ConfirmStatus != "N")
                                {
                                    throw new SystemException("進貨單頭或單身狀態已核單，無法修改!!");
                                }

                                if (ReceiptQty != item.ReceiptQty)
                                {
                                    throw new SystemException("【檢驗單進貨數量】與【進貨單進貨數量】不同!!");
                                }

                                if (AcceptQty + ReturnQty != ReceiptQty)
                                {
                                    throw new SystemException("【驗收數量】+【驗退數量】需等於【進貨數量】!!");
                                }

                                GrDetailId = item.GrDetailId;
                                GrId = item.GrId;
                                OrigUnitPrice = item.OrigUnitPrice;
                            }
                            #endregion

                            #region //UPDATE 進貨單相關資料

                            #region //UPDATE SCM.GrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GrDetail SET
                                    AcceptanceDate = @AcceptanceDate,
                                    AcceptQty = @AcceptQty,
                                    AvailableQty = @AcceptQty,
                                    ReturnQty = @ReturnQty,
                                    QcStatus = @QcStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                AcceptanceDate,
                                AcceptQty,
                                ReturnQty,
                                QcStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrDetailId
                            });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(erpHelper.UpdateGrTotal(GrId, -1, sqlConnection, sqlConnection2, CurrentUser));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
                            }
                            #endregion

                            #region //拋轉ERP
                            JObject updateTransferGoodsReceiptResult = JObject.Parse(erpHelper.UpdateTransferGoodsReceipt(GrId, sqlConnection, sqlConnection2, CurrentUser, CurrentCompany));
                            if (updateTransferGoodsReceiptResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateTransferGoodsReceiptResult["msg"].ToString());
                            }
                            #endregion

                            #endregion

                            #region //確認是否為急料，若為急料則由品保直接核單
                            if (QuickStatus == "Y")
                            {
                                #region //確認是否有附件
                                if (QcGoodsReceiptLogFile.Length <= 0)
                                {
                                    throw new SystemException("若為急料，需檢附證明文件!!");
                                }
                                #endregion

                                JObject confirmGrDetailResult = JObject.Parse(erpHelper.ConfirmGrDetail(GrDetailId, sqlConnection, sqlConnection2, CurrentUser));
                                if (confirmGrDetailResult["status"].ToString() != "success")
                                {
                                    throw new SystemException(confirmGrDetailResult["msg"].ToString());
                                }
                            }
                            #endregion

                            #region //UPDATE MES.QcGoodsReceiptLog
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcGoodsReceiptLog SET
                                    AcceptQty = @AcceptQty,
                                    ReturnQty = @ReturnQty,
                                    AcceptanceDate = @AcceptanceDate,
                                    QcStatus = @QcStatus,
                                    QuickStatus = @QuickStatus,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE LogId = @LogId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                AcceptQty,
                                ReturnQty,
                                AcceptanceDate,
                                QcStatus,
                                QuickStatus,
                                Remark,
                                LastModifiedDate,
                                LastModifiedBy,
                                LogId
                            });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INSERT MES.QcGoodsReceiptLogFile
                            #region //刪除原本檔案
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.QcGoodsReceiptLogFile
                                    WHERE LogId = @LogId";
                            dynamicParameters.Add("LogId", LogId);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #endregion

                            if (QcGoodsReceiptLogFile.Length > 0)
                            {
                                var qcGoodsReceiptLogFileList = QcGoodsReceiptLogFile.Split(',');

                                foreach (var fileId in qcGoodsReceiptLogFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcGoodsReceiptLogFile (LogId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@LogId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LogId,
                                            FileId = fileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                }
                            }
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

        #region //CancelMeasurementUpload -- 取消進貨檢驗單據上傳 -- Ann 2025-02-10
        public string CancelMeasurementUpload(int QcRecordId)
        {
            int rowsAffected = 0;

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷確認量測單據資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CheckQcMeasureData
                                , b.QcGoodsReceiptId, b.GrDetailId
                                , c.ReceiptQty, c.ConfirmStatus
                                , d.MeasureType
                                FROM MES.QcRecord a 
                                INNER JOIN MES.QcGoodsReceipt b ON a.QcRecordId = b.QcRecordId
                                INNER JOIN SCM.GrDetail c ON b.GrDetailId = c.GrDetailId
                                INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測單據資料錯誤!");

                        int QcGoodsReceiptId = -1;
                        int GrDetailId = -1;
                        string QcStatus = "";
                        int ReceiptQty = 0;
                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "Y" && item.CheckQcMeasureData != "P")
                            {
                                throw new SystemException("此量測單據尚未上傳，無法反確認!!");
                            }

                            if (item.ConfirmStatus != "N")
                            {
                                throw new SystemException("進貨單非未確認狀態，量測數據無法反確認!!");
                            }

                            QcGoodsReceiptId = item.QcGoodsReceiptId;
                            GrDetailId = item.GrDetailId;
                            ReceiptQty = item.ReceiptQty;

                            if (item.MeasureType == "0")
                            {
                                QcStatus = item.MeasureType;
                            }
                            else
                            {
                                QcStatus = "1";
                            }
                        }
                        #endregion

                        #region //UPDATE MES.QcRecord CheckQcMeasureData
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                CheckQcMeasureData = 'N',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Update MES.QcGoodsReceipt
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcGoodsReceipt SET
                                QcStatus = 'A',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcGoodsReceiptId = @QcGoodsReceiptId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                QcGoodsReceiptId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Update SCM.GrDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                AcceptQty = 0,
                                AvailableQty = 0,
                                ReturnQty = 0,
                                QcStatus = @QcStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrDetailId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除上傳數據
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.QcMeasureData 
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除品異單
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM QMS.AqBarcode a
                                INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                WHERE b.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.Abnormalquality
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除檢驗作業單據
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM MES.QcGoodsReceiptLogFile a 
                                INNER JOIN MES.QcGoodsReceiptLog b ON a.LogId = b.LogId
                                WHERE b.QcGoodsReceiptId = @QcGoodsReceiptId";
                        dynamicParameters.Add("QcGoodsReceiptId", QcGoodsReceiptId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.QcGoodsReceiptLog
                                WHERE QcGoodsReceiptId = @QcGoodsReceiptId";
                        dynamicParameters.Add("QcGoodsReceiptId", QcGoodsReceiptId);

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

        #region //UpdatedOutsourcingQcRecord -- 更新量測記錄單頭資料  -- GPAI 2025-04-10
        public string UpdatedOutsourcingQcRecord(int QcRecordId, int QcNoticeId, int QcTypeId, int MtlItemId, string Remark, int CurrentFileId, string QcRecordFile, string InputType, string ServerPath, string ServerPath2, string SupportAqFlag, string ResolveFile, string ResolveFileJson, string QcRecordFileByNas)
        {
            try
            {
                if (InputType.Length < 0) throw new SystemException("【輸入方式】不能為空!");
                if (SupportAqFlag.Length < 0) throw new SystemException("【是否支援品異單】不能為空!");
                //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Cavity") throw new SystemException("【穴號模式】只能由全吋檢或成型工程檢使用!!");
                //if (QcType != "PVTQC" && QcType != "TQC" && InputType == "Picture") throw new SystemException("【面型圖模式】只能由全吋檢或成型工程檢使用!!");

                string ErpDbName = "";
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //檢查此量測資料是否已經有上傳數據
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcMeasureData
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcMeasureDataResult.Count() > 0) throw new SystemException("此量測資料已有上傳數據，無法更改!");
                            #endregion

                            #region //判斷量測記錄資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CheckQcMeasureData
                                    FROM MES.QcRecord
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");

                            string CheckQcMeasureData = "";
                            foreach (var item in QcRecordResult)
                            {
                                if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "S") throw new SystemException("量測單據狀態無法更改!!");
                                CheckQcMeasureData = item.CheckQcMeasureData;
                            }
                            #endregion

                            #region //判斷MES製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * 
                                    FROM PDM.MtlItem 
                                    WHERE MtlItemId  = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("MES品號資料錯誤!");

                            string MtlItemNo = "";
                            string MtlItemName = "";
                            foreach (var item in result)
                            {
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                            }
                            #endregion

                            #region //判斷ERP製令狀態是否正確
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT TA011, TA013
                            //        FROM MOCTA
                            //        WHERE TA001 = @TA001
                            //        AND TA002 = @TA002";
                            //dynamicParameters.Add("TA001", WoErpPrefix);
                            //dynamicParameters.Add("TA002", WoErpNo);

                            //var MOCTAResult = sqlConnection2.Query(sql, dynamicParameters);

                            //if (MOCTAResult.Count() <= 0) throw new SystemException("ERP製令資料錯誤!!");

                            //foreach (var item in MOCTAResult)
                            //{
                            //    if (SupportAqFlag == "Y" && (item.TA011 == "Y" || item.TA011 == "y" || item.TA013 == "V")) throw new SystemException("ERP製令狀態無法開立!!");
                            //}
                            #endregion

                            #region //判斷量測類型資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                    FROM QMS.QcType a 
                                    WHERE a.QcTypeId = @QcTypeId";
                            dynamicParameters.Add("QcTypeId", QcTypeId);

                            var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                            string QcTypeNo = "";
                            string QcTypeName = "";
                            foreach (var item in QcTypeResult)
                            {
                                if (item.SupportProcessFlag == "Y" /*&& MoProcessId <= 0*/) throw new SystemException("【量測製程】不能為空!");
                                QcTypeNo = item.QcTypeNo;
                                QcTypeName = item.QcTypeName;
                            }
                            #endregion

                            //if (MoProcessId > 0)
                            //{
                            //    #region //判斷製程資料是否正確
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"SELECT TOP 1 1
                            //            FROM MES.MoProcess
                            //            WHERE MoProcessId = @MoProcessId";
                            //    dynamicParameters.Add("MoProcessId", MoProcessId);

                            //    var MoProcessResult = sqlConnection.Query(sql, dynamicParameters);
                            //    if (MoProcessResult.Count() <= 0) throw new SystemException("製令製程資料錯誤!");
                            //    #endregion
                            //}

                            #region //建立CheckQcMeasureData
                            //if (InputType == "Cavity" || InputType == "Picture")
                            //{
                            //    if (QcRecordFile.Length > 0) CheckQcMeasureData = "Y";
                            //}
                            #endregion

                            #region //UPDATE MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    QcNoticeId = @QcNoticeId,
                                    QcTypeId = @QcTypeId,
                                    InputType = @InputType,
                                    Remark = @Remark,
                                    CurrentFileId = @CurrentFileId,
                                    CheckQcMeasureData = @CheckQcMeasureData,
                                    SupportAqFlag = @SupportAqFlag,
                                    //ResolveFileJson = @ResolveFileJson,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcNoticeId = QcNoticeId <= 0 ? (int?)null : QcNoticeId,
                                    QcTypeId,
                                    InputType,
                                    //MoId,
                                    //MoProcessId = MoProcessId <= 0 ? (int?)null : MoProcessId,
                                    Remark,
                                    CurrentFileId,
                                    CheckQcMeasureData,
                                    SupportAqFlag,
                                    ResolveFileJson,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新FILE資訊

                            #region //先將原本資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.QcRecordFile
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", QcRecordId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INSERT MES.QcRecordFile
                            if (QcRecordFile.Length > 0)
                            {
                                var qcRecordFileList = QcRecordFile.Split(',');

                                foreach (var qcRecordFileId in qcRecordFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileId = qcRecordFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT MES.QcRecordFile By NAS
                            if (QcRecordFileByNas.Length > 0)
                            {
                                JObject uploadFileJson = JObject.Parse(QcRecordFileByNas);

                                foreach (var item in uploadFileJson["uploadFileInfo"])
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
                                            FileType = "other",
                                            PhysicalPath = item["FilePath"].ToString(),
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion

                            #region //INSERT ResolveFile
                            if (ResolveFile.Length > 0)
                            {
                                var resolveFileList = ResolveFile.Split(',');

                                foreach (var resolveFileId in resolveFileList)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecordFile (QcRecordId, FileType, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordFileId, INSERTED.QcRecordId, INSERTED.FileId
                                            VALUES (@QcRecordId, @FileType, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            FileType = "resolve",
                                            FileId = resolveFileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                            }
                            #endregion
                            #endregion

                            #region //將量測EXCEL先存回實體路徑並重新讀取
                            string FileName = "";
                            if (CurrentFileId > 0)
                            {
                                #region //取得原本加密檔案資料，並重新上傳
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId, a.FileContent, a.FileName, a.FileExtension, a.FileSize, a.ClientIP, a.Source
                                        FROM BAS.[File] a
                                        WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", CurrentFileId);

                                var FileResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in FileResult)
                                {
                                    FileName = item.FileName;
                                    ServerPath = Path.Combine(ServerPath, item.FileName + "-" + item.FileId + item.FileExtension);
                                    byte[] fileContent = (byte[])item.FileContent;
                                    File.WriteAllBytes(ServerPath, fileContent); // Requires System.IO

                                    #region //將檔案重新上傳
                                    System.Net.WebClient wc = new System.Net.WebClient();
                                    byte[] newFilebytes = wc.DownloadData(ServerPath);

                                    dynamicParameters = new DynamicParameters();
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
                                            CompanyId = CurrentCompany,
                                            item.FileName,
                                            FileContent = newFilebytes,
                                            item.FileExtension,
                                            item.FileSize,
                                            item.ClientIP,
                                            item.Source,
                                            DeleteStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    int newFileId = -1;
                                    foreach (var item2 in insertResult)
                                    {
                                        newFileId = item2.FileId;
                                    }
                                    #endregion

                                    #region //UPDATE原本FileId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcRecord SET
                                            CurrentFileId = @CurrentFileId,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CurrentFileId = newFileId,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QcRecordId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除實體檔案
                                    File.Delete(ServerPath);
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            if ((CheckQcMeasureData == "Y" || CheckQcMeasureData == "P") && CurrentFileId > 0 && (QcTypeNo == "PVTQC" || QcTypeNo == "TQC"))
                            {
                                //SendQcMail(sqlConnection, QcRecordId, "", MtlItemNo, MtlItemName, QcTypeNo, QcTypeName
                                //    , Remark, FileName, ServerPath2, -1);
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

        #region //UpdateCheckGoodsOutSourcing -- 確認收貨狀態 -- GPAI 2025-04-17
        public string UpdateCheckGoodsOutSourcing(int QcRecordId, string CheckQcMeasureData, string Disallowance)
        {
            try
            {
                if (CheckQcMeasureData.Length < 0) throw new SystemException("【更改狀態】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo
                                FROM BAS.Company a 
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //取得收件人員資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo, a.UserName
                                FROM BAS.[User] a 
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var CheckUserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CheckUserResult.Count() <= 0) throw new SystemException("人員資料錯誤!!");

                        string UserNo = "";
                        string UserName = "";
                        foreach (var item in CheckUserResult)
                        {
                            UserNo = item.UserNo;
                            UserName = item.UserName;
                        }
                        #endregion

                        #region //判斷量測記錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId, a.CheckQcMeasureData, ISNULL(FORMAT(a.ReceiptDate, 'yyyy-MM-dd'), '') ReceiptDate
                                , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') DocCreateDate
                                , e.DepartmentId
                                , f.DepartmentNo, f.DepartmentName
                                , g.QcTypeNo, g.QcTypeName
                                , h.MtlItemNo, h.MtlItemName
                                FROM MES.QcRecord a 
                                INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId
                                INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                                INNER JOIN QMS.QcType g ON a.QcTypeId = g.QcTypeId
                                INNER JOIN PDM.MtlItem h ON a.MtlItemId = h.MtlItemId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");

                        string WoErpPrefix = "";
                        string WoErpNo = "";
                        string WoSeq = "";
                        string QcTypeNo = "";
                        string QcTypeName = "";
                        string ProcessAlias = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ReceiptDate = "";
                        string MtlItemNo = "";
                        string MtlItemName = "";
                        int DepartmentId = -1;
                        int ModeId = -1;
                        string ModeName = "";
                        string DocCreateDate = "";
                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData == "Y") throw new SystemException("單據已上傳完成!!");
                            if (item.CheckQcMeasureData == "S") throw new SystemException("單據已開始量測!!");
                            QcTypeNo = item.QcTypeNo;
                            QcTypeName = item.QcTypeName;
                            DepartmentNo = item.DepartmentNo;
                            DepartmentName = item.DepartmentName;
                            ReceiptDate = item.ReceiptDate.ToString();
                            MtlItemNo = item.MtlItemNo;
                            MtlItemName = item.MtlItemName;
                            DepartmentId = item.DepartmentId;
                            DocCreateDate = item.DocCreateDate.ToString();
                        }
                        #endregion

                        #region //確認收貨及開始量測欄位
                        string QcStartDate = "";
                        string CurrentDatetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        switch (CheckQcMeasureData)
                        {
                            case "C":
                                ReceiptDate = CurrentDatetime;
                                break;
                            case "S":
                                QcStartDate = CurrentDatetime;
                                if (ReceiptDate == "")
                                {
                                    ReceiptDate = QcStartDate;
                                }
                                break;
                            case "F":
                                ReceiptDate = CurrentDatetime;
                                break;
                            default:
                                throw new SystemException("單據狀態錯誤!!");
                        }
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                CheckQcMeasureData = @CheckQcMeasureData,
                                ReceiptDate = @ReceiptDate,
                                QcStartDate = @QcStartDate,
                                DisallowanceReason = @Disallowance,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CheckQcMeasureData,
                                ReceiptDate,
                                QcStartDate,
                                Disallowance,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //MAMO推播通知
                        //string Content = "";

                        //#region //取得送測條碼清單
                        //string barcodeInfoDesc = "| 條碼 | 刻字 |\n| :-- | :-- |\n";
                        //int count = 0;

                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT a.BarcodeId
                        //        , ISNULL(c.ItemValue, '') ItemValue
                        //        , d.BarcodeNo
                        //        FROM MES.QcReceiptBarcode a 
                        //        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                        //        LEFT JOIN MES.BarcodeAttribute c ON a.BarcodeId = c.BarcodeId AND c.MoId = b.MoId AND c.ItemNo = 'Lettering'
                        //        INNER JOIN MES.Barcode d ON a.BarcodeId = d.BarcodeId
                        //        WHERE a.QcRecordId = @QcRecordId";
                        //dynamicParameters.Add("QcRecordId", QcRecordId);

                        //var QcReceiptBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        //foreach (var item in QcReceiptBarcodeResult)
                        //{
                        //    barcodeInfoDesc += "| " + item.BarcodeNo + " | " + item.ItemValue + " |\n";
                        //    count++;
                        //}
                        //#endregion

                        //if (CheckQcMeasureData == "C")
                        //{
                        //    Content = "### 【量測確認收件通知】\n" +
                        //                    "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                        //                    "##### 品號: " + MtlItemNo + "\n" +
                        //                    "##### 品名: " + MtlItemName + "\n" +
                        //                    "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                        //                    "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                        //                    "- 製程: " + ProcessAlias + "\n" +
                        //                    "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                        //                    "- 送檢時間: " + DocCreateDate + "\n" +
                        //                    "- 收件人員: " + UserNo + " " + UserName + "\n" +
                        //                    "- 確認收件時間: " + CreateDate.ToString() + "\n" +
                        //                    "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        //}
                        //else if (CheckQcMeasureData == "F")
                        //{
                        //    Content = "### 【量測單據駁回通知】\n" +
                        //                    "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                        //                    "##### 品號: " + MtlItemNo + "\n" +
                        //                    "##### 品名: " + MtlItemName + "\n" +
                        //                    "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                        //                    "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                        //                    "- 製程: " + ProcessAlias + "\n" +
                        //                    "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                        //                    "- 送檢時間: " + DocCreateDate + "\n" +
                        //                    "- 駁回人員: " + UserNo + " " + UserName + "\n" +
                        //                    "- 駁回時間: " + CreateDate.ToString() + "\n" +
                        //                    "- 駁回原因: " + Disallowance + "\n" +
                        //                    "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        //}
                        //else if (CheckQcMeasureData == "S")
                        //{
                        //    Content = "### 【開始量測通知】\n" +
                        //                    "##### 製令: " + WoErpPrefix + "-" + WoErpNo + "(" + WoSeq + ")\n" +
                        //                    "##### 品號: " + MtlItemNo + "\n" +
                        //                    "##### 品名: " + MtlItemName + "\n" +
                        //                    "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                        //                    "- 量測單據編號: " + QcRecordId.ToString() + "\n" +
                        //                    "- 製程: " + ProcessAlias + "\n" +
                        //                    "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                        //                    "- 送檢時間: " + DocCreateDate + "\n" +
                        //                    "- 量測人員: " + UserNo + " " + UserName + "\n" +
                        //                    "- 開始量測時間: " + CreateDate.ToString() + "\n" +
                        //                    "- 送檢條碼:\n" + barcodeInfoDesc + "\n";
                        //}
                        //else
                        //{
                        //    throw new SystemException("確認收貨類型錯誤!!");
                        //}

                        //#region //確認推播群組
                        //string SendId = "";
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT a.ChannelId
                        //        FROM MES.QcMamoChannel a
                        //        WHERE a.ModeId = @ModeId
                        //        AND PushType = @PushType";
                        //dynamicParameters.Add("ModeId", 39);
                        //dynamicParameters.Add("PushType", CheckQcMeasureData);

                        //var QcMamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                        //if (QcMamoChannelResult.Count() <= 0) throw new SystemException("收貨確認/駁回推播群組資料錯誤!!");

                        //int ChannelId = -1;
                        //foreach (var item in QcMamoChannelResult)
                        //{
                        //    ChannelId = item.ChannelId;
                        //    SendId = ChannelId.ToString();
                        //}
                        //#endregion

                        //#region //取得標記USER資料(原送測人員部門)
                        //List<string> Tags = new List<string>();
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT a.UserId
                        //        , b.UserNo, b.UserName
                        //        FROM MAMO.ChannelMembers a 
                        //        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        //        WHERE a.ChannelId = @SendId
                        //        AND b.DepartmentId = @DepartmentId";
                        //dynamicParameters.Add("SendId", SendId);
                        //dynamicParameters.Add("DepartmentId", DepartmentId);

                        //var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        //foreach (var item in UserResult)
                        //{
                        //    Tags.Add(item.UserNo);
                        //}
                        //#endregion

                        //List<int> Files = new List<int>();

                        //string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", SendId, Content, Tags, Files);

                        //JObject MamoResultJson = JObject.Parse(MamoResult);
                        //if (MamoResultJson["status"].ToString() != "success")
                        //{
                        //    throw new SystemException(MamoResultJson["msg"].ToString());
                        //}
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
        #region //DeleteQcRecord -- 刪除量測記錄資料 -- Ann 2022-04-13
        public string DeleteQcRecord(int QcRecordId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CreateBy, a.CheckQcMeasureData
                                , b.UserNo, b.UserName
                                FROM MES.QcRecord a
                                INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N") throw new SystemException("單據狀態不可刪除!!");
                            //if (item.CreateBy != CreateBy) throw new SystemException("此單據由【" + item.UserNo + "-" + item.UserName + "】開立，無法刪除!!");
                        }
                        #endregion

                        #region //判斷此量測記錄是否已經有上傳歷程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcMeasureDataResult.Count() > 0) throw new SystemException("此單據已經有上傳數據紀錄，無法刪除!!");
                        #endregion

                        #region //刪除MES單據

                        #region //刪除檔案TABLE
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcRecordFile
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除暫存TABLE
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcMeasureDataTemp
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcRecord
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

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

        #region //DeleteTempSpreadsheet -- 刪除暫存量測記錄 -- Ann 2023-06-15
        public string DeleteTempSpreadsheet(int QcRecordId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CreateBy, a.CheckQcMeasureData
                                , b.UserNo, b.UserName
                                FROM MES.QcRecord a
                                INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "S") throw new SystemException("此單據已有上傳紀錄，無法刪除!!");
                            //if (item.CreateBy != CreateBy) throw new SystemException("此單據由【" + item.UserNo + "-" + item.UserName + "】開立，無法刪除!!");
                        }
                        #endregion

                        #region //判斷此量測記錄是否已經有上傳歷程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcMeasureDataResult.Count() > 0) throw new SystemException("此單據已經有上傳數據紀錄，無法刪除!!");
                        #endregion

                        #region //更新MES.QcRecord SpreadsheetData
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                SpreadsheetData = null,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //DeleteResolveFileJson -- 刪除暫存量測解析JSON -- Ann 2023-11-15
        public string DeleteResolveFileJson(int QcRecordId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CreateBy, a.CheckQcMeasureData
                                , b.UserNo, b.UserName
                                FROM MES.QcRecord a
                                INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "S") throw new SystemException("此單據已有上傳紀錄，無法刪除!!");
                        }
                        #endregion

                        #region //判斷此量測記錄是否已經有上傳歷程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcMeasureDataResult.Count() > 0) throw new SystemException("此單據已經有上傳數據紀錄，無法刪除!!");
                        #endregion

                        #region //更新MES.QcRecord SpreadsheetData
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                ResolveFileJson = null,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //DeleteQcGoodsReceipt -- 刪除進貨檢量測單據 -- Ann 2024-04-08
        public string DeleteQcGoodsReceipt(int QcRecordId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CreateBy, a.CheckQcMeasureData
                                , b.UserNo, b.UserName
                                FROM MES.QcRecord a
                                INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        foreach (var item in QcRecordResult)
                        {
                            if (item.CheckQcMeasureData != "N") throw new SystemException("單據狀態不可刪除!!");
                            //if (item.CreateBy != CreateBy) throw new SystemException("此單據由【" + item.UserNo + "-" + item.UserName + "】開立，無法刪除!!");
                        }
                        #endregion

                        #region //判斷此量測記錄是否已經有上傳歷程
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcMeasureDataResult.Count() > 0) throw new SystemException("此單據已經有上傳數據紀錄，無法刪除!!");
                        #endregion

                        #region //刪除MES單據

                        #region //刪除檔案TABLE
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcRecordFile
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除進貨檢table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcGoodsReceipt
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcRecord
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

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

        #region //DeleteQcRecordFile -- 刪除量測歸檔資料 -- Ann 2022-04-13
        public string DeleteQcRecordFile(int QcRecordFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.QcRecordFile a 
                                WHERE a.QcRecordFileId = @QcRecordFileId
                                AND a.FileType = 'file-management'";
                        dynamicParameters.Add("QcRecordFileId", QcRecordFileId);

                        var QcRecordFileResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordFileResult.Count() <= 0) throw new SystemException("量測歸檔資料錯誤!!");
                        #endregion

                        #region //刪除檔案TABLE
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.QcRecordFile
                                WHERE QcRecordFileId = @QcRecordFileId";
                        dynamicParameters.Add("QcRecordFileId", QcRecordFileId);

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

        #region //DeleteQcMachineModePlanning -- 刪除量測機台排程資料 -- Ann 2024-08-22
        public string DeleteQcMachineModePlanning(int QmmpId)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測機台排程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Sort, a.QmmDetailId, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                                , b.QcRecordId, b.ConfirmStatus
                                , c.CheckQcMeasureData
                                FROM QMS.QcMachineModePlanning a 
                                INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                INNER JOIN MES.QcRecord c ON b.QcRecordId = c.QcRecordId
                                WHERE a.QmmpId = @QmmpId";
                        dynamicParameters.Add("QmmpId", QmmpId);

                        var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMachineModePlanningResult.Count() <= 0) throw new SystemException("量測機台排程資料錯誤!!");

                        int Sort = -1;
                        int QmmDetailId = -1;
                        int QcRecordId = -1;
                        foreach (var item in QcMachineModePlanningResult)
                        {
                            if (item.ConfirmStatus != "N")
                            {
                                throw new SystemException("量測單據排程已確認，無法刪除!!");
                            }

                            if (item.CheckQcMeasureData != "N" && item.CheckQcMeasureData != "A" && item.CheckQcMeasureData != "C")
                            {
                                throw new SystemException("量測單據狀態錯誤!!");
                            }

                            if (item.CreateDate != DateTime.Now.ToString("yyyy-MM-dd"))
                            {
                                throw new SystemException("不能夠異動非當日之排程資料!!");
                            }

                            Sort = item.Sort;
                            QmmDetailId = item.QmmDetailId;
                            QcRecordId = item.QcRecordId;
                        }
                        #endregion

                        #region //排程計算核心邏輯

                        #region //取得刪除後新排程初始開始日期
                        DateTime firstStartDate = DateTime.Now;
                        if (Sort != 1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.EstimatedEndDate
                                    FROM QMS.QcMachineModePlanning a 
                                    WHERE a.Sort = @Sort
                                    AND QmmDetailId = @QmmDetailId
                                    AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow";
                            dynamicParameters.Add("Sort", Sort - 1);
                            dynamicParameters.Add("QmmDetailId", QmmDetailId);
                            dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                            var EstimatedEndDateResult = sqlConnection.Query(sql, dynamicParameters);

                            if (EstimatedEndDateResult.Count() <= 0) throw new SystemException("取得初始開始日期時錯誤!!");

                            foreach (var item in EstimatedEndDateResult)
                            {
                                firstStartDate = item.EstimatedEndDate;
                            }
                        }
                        #endregion

                        #region //取得所有異動排程資料
                        List<int> OriQmmpIdSortArray = new List<int>();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QmmpId, a.Sort, a.EstimatedStartDate, a.EstimatedEndDate
                                , b.EstimatedMeasurementTime
                                , c.QcStartDate, c.QcRecordId
                                FROM QMS.QcMachineModePlanning a 
                                INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                INNER JOIN MES.QcRecord c ON b.QcRecordId = c.QcRecordId
                                WHERE a.Sort > @Sort
                                AND a.QmmDetailId = @QmmDetailId
                                AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow
                                ORDER BY a.Sort";
                        dynamicParameters.Add("Sort", Sort);
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);
                        dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                        var PlanningSortCheckResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in PlanningSortCheckResult)
                        {
                            OriQmmpIdSortArray.Add(item.QmmpId);
                        }
                        #endregion

                        #region //Remove原先排程順序
                        if (OriQmmpIdSortArray.Count() > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                    Sort = NULL,
                                    EstimatedStartDate = NULL,
                                    EstimatedEndDate = NULL
                                    WHERE QmmpId IN (" + string.Join(",", OriQmmpIdSortArray) + ")";

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //先刪除此次排程
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcMachineModePlanning
                                WHERE QmmpId = @QmmpId";
                        dynamicParameters.Add("QmmpId", QmmpId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Reset原先排程順序
                        int i = 0;
                        DateTime currentEstimatedEndDate = firstStartDate;
                        foreach (var qmmpId in OriQmmpIdSortArray)
                        {
                            #region //取得此排程相關資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.QcRecordId
                                    FROM QMS.QcMachineModePlanning a 
                                    INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                    WHERE a.QmmpId = @QmmpId";
                            dynamicParameters.Add("QmmpId", qmmpId);

                            var QcRecordIdResult = sqlConnection.Query(sql, dynamicParameters);

                            int currentQcRecordId = 0;
                            foreach (var item in QcRecordIdResult)
                            {
                                currentQcRecordId = item.QcRecordId;
                            }
                            #endregion

                            #region //確認此量測單據是否已經有其他機型已排定排程，若有則需檢核時間順序
                            if (i == 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.QmmDetailId, a.EstimatedEndDate
                                        , d.MachineDesc
                                        FROM QMS.QcMachineModePlanning a 
                                        INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                        INNER JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                        INNER JOIN MES.Machine d ON c.MachineId = d.MachineId
                                        WHERE b.QcRecordId = @QcRecordId
                                        AND a.QmmpId != @QmmpId
                                        AND FORMAT(a.CreateDate, 'yyyy-MM-dd') = @DateNow
                                        ORDER BY a.EstimatedEndDate DESC";
                                dynamicParameters.Add("QcRecordId", currentQcRecordId);
                                dynamicParameters.Add("QmmpId", qmmpId);
                                dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));

                                var QcMachineModePlanningResult2 = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in QcMachineModePlanningResult2)
                                {
                                    DateTime finalEndDate = item.EstimatedEndDate;
                                    if (currentEstimatedEndDate < finalEndDate)
                                    {
                                        currentEstimatedEndDate = finalEndDate;
                                        //throw new SystemException("此次量測單號【" + QcRecordId + "】排程開始時間【" + firstStartDate + "】不可早於上個排程【" + item.MachineDesc + "】結束時間【" + finalEndDate.ToString("yyyy-MM-dd HH:mm:ss") + "】!!");
                                    }
                                }
                            }
                            #endregion

                            #region //計算新的開始日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.EstimatedMeasurementTime
                                    FROM QMS.QcMachineModePlanning a 
                                    INNER JOIN QMS.QcRecordPlanning b ON a.QrpId = b.QrpId
                                    WHERE a.QmmpId = @QmmpId";
                            dynamicParameters.Add("QmmpId", qmmpId);

                            var GetEstimatedMeasurementTimeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GetEstimatedMeasurementTimeResult.Count() <= 0) throw new SystemException("【量測機台排程ID: " + qmmpId + "】查無此排程資料!!");

                            double currentEstimatedMeasurementTime = 0;
                            foreach (var item in GetEstimatedMeasurementTimeResult)
                            {
                                currentEstimatedMeasurementTime = item.EstimatedMeasurementTime;
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.QcMachineModePlanning SET 
                                    Sort = @Sort,
                                    EstimatedStartDate = @EstimatedStartDate,
                                    EstimatedEndDate = @EstimatedEndDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QmmpId = @QmmpId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Sort = Sort + i,
                                    EstimatedStartDate = currentEstimatedEndDate,
                                    EstimatedEndDate = currentEstimatedEndDate.AddMinutes(currentEstimatedMeasurementTime),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QmmpId = qmmpId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            currentEstimatedEndDate = currentEstimatedEndDate.AddMinutes(currentEstimatedMeasurementTime);

                            i++;
                        }
                        #endregion

                        #endregion

                        #region //更改量測單據排程狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET 
                                MeasurementPlanning = 'N',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //Upload Data
        #region //UploadQcData -- 解析SpreadSheet Data及上傳量測數據 -- Ann 2023-03-30
        public string UploadQcData(int QcRecordId, string QcData, string BarcodeQcRecordData, string SpreadsheetData, string ServerPath2)
        {
            try
            {
                if (QcData.Length <= 0) throw new SystemException("尚未接收到數據!!");
                if (BarcodeQcRecordData.Length <= 0) throw new SystemException("尚未接收到數據!!");
                if (SpreadsheetData.Length <= 0) throw new SystemException("尚未接收完整Spreadsheet數據!!");

                int rowsAffected = 0;
                int? BarcodeId = -1;
                int BarcodeQty = -1;
                List<int?> BarcodeIdList = new List<int?>();
                Dictionary<int, int?> QcBarcodeList = new Dictionary<int, int?>();

                List<QcNgCode> QcNgCodes = new List<QcNgCode>();
                string QcCauseJsonString = "";
                int MoId = -1;
                int QcTypeId = -1;
                int? MoProcessId = -1;
                string ProcessCheckType = "";
                string CheckQcMeasureData = "";
                string checkEditFlag = "N";
                string Remark = "";
                bool passBool = true;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否有誤
                        sql = @"SELECT a.CompanyNo
                                FROM BAS.Company a 
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //檢查量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoId, a.QcTypeId, a.MoProcessId, a.CheckQcMeasureData, a.Remark, ISNULL(a.CurrentFileId, -1) CurrentFileId, a.SupportAqFlag
                                , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') ReceiptDate
                                , ISNULL(b.FileName, '') FileName
                                , ISNULL(c.ProcessAlias, '') ProcessAlias
                                , e.DepartmentId
                                , f.DepartmentNo, f.DepartmentName
                                FROM MES.QcRecord a
                                LEFT JOIN BAS.[File] b ON a.CurrentFileId = b.FileId
                                LEFT JOIN MES.MoProcess c ON a.MoProcessId = c.MoProcessId
                                INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId
                                INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        string FileName = "";
                        int CurrentFileId = -1;
                        string SupportAqFlag = "";
                        string ProcessAlias = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ReceiptDate = "";
                        int DepartmentId = -1;
                        string OriCheckQcMeasureData = "";
                        foreach (var item2 in QcRecordResult)
                        {
                            if (item2.CheckQcMeasureData == "A") throw new SystemException("單據尚未確認收件，無法上傳!!");
                            if (item2.CheckQcMeasureData == "C") throw new SystemException("單據尚未開始量測，無法上傳!!");

                            MoId = item2.MoId != null ? item2.MoId : -1;
                            QcTypeId = item2.QcTypeId;
                            MoProcessId = item2.MoProcessId;
                            CheckQcMeasureData = item2.CheckQcMeasureData == "S" ? "N" : item2.CheckQcMeasureData;
                            Remark = item2.Remark;
                            FileName = item2.FileName;
                            CurrentFileId = item2.CurrentFileId;
                            SupportAqFlag = item2.SupportAqFlag;
                            ProcessAlias = item2.ProcessAlias;
                            DepartmentNo = item2.DepartmentNo;
                            DepartmentName = item2.DepartmentName;
                            ReceiptDate = item2.ReceiptDate.ToString();
                            DepartmentId = item2.DepartmentId;
                            OriCheckQcMeasureData = item2.CheckQcMeasureData;
                        }
                        #endregion

                        #region //判斷MES製令資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModeId
                                , (b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(varchar(10), a.WoSeq) + ')') WoErpFullNo
                                , c.MtlItemNo, c.MtlItemName
                                , d.ModeName
                                FROM MES.ManufactureOrder a 
                                INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                INNER JOIN MES.ProdMode d ON a.ModeId = d.ModeId
                                WHERE MoId = @MoId";
                        dynamicParameters.Add("MoId", MoId);

                        var ManufactureOrderResult = sqlConnection.Query(sql, dynamicParameters);
                        if (ManufactureOrderResult.Count() <= 0) throw new SystemException("MES製令資料錯誤!");

                        string WoErpFullNo = "";
                        string MtlItemNo = "";
                        string MtlItemName = "";
                        int ModeId = -1;
                        string ModeName = "";
                        foreach (var item in ManufactureOrderResult)
                        {
                            WoErpFullNo = item.WoErpFullNo;
                            MtlItemNo = item.MtlItemNo;
                            MtlItemName = item.MtlItemName;
                            ModeId = item.ModeId;
                            ModeName = item.ModeName;
                        }
                        #endregion

                        #region //判斷量測類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                    FROM QMS.QcType a 
                                    WHERE a.QcTypeId = @QcTypeId";
                        dynamicParameters.Add("QcTypeId", QcTypeId);

                        var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                        string QcTypeNo = "";
                        string QcTypeName = "";
                        string SupportProcessFlag = "";
                        foreach (var item in QcTypeResult)
                        {
                            if (item.SupportProcessFlag == "Y" && MoProcessId <= 0) throw new SystemException("【量測製程】不能為空!");
                            QcTypeNo = item.QcTypeNo;
                            QcTypeName = item.QcTypeName;
                            SupportProcessFlag = item.SupportProcessFlag;
                        }
                        #endregion

                        #region //確認是否已上傳過量測數據
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMeasureDataResult.Count() > 0 && CheckQcMeasureData != "P") throw new SystemException("此量測記錄單已上傳過量測記錄，無法重複上傳!!");
                        #endregion

                        #region //判斷製令製程資料是否正確
                        if (MoProcessId != null)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ProcessCheckType
                                FROM MES.MoProcess
                                WHERE MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            var result12 = sqlConnection.Query(sql, dynamicParameters);
                            if (result12.Count() <= 0) throw new SystemException("製令製程資料錯誤!");

                            foreach (var item2 in result12)
                            {
                                ProcessCheckType = item2.ProcessCheckType;
                            }
                        }
                        #endregion

                        #region //解析BarcodeQcRecordData
                        var barcodeQcRecordJson = JObject.Parse(BarcodeQcRecordData);

                        List<string> BarcodeList = new List<string>();
                        List<string> SpreadSheetBarcodeList = new List<string>();

                        foreach (var item in barcodeQcRecordJson["barcodeQcRecordInfo"])
                        {
                            if (item["QcStatus"].ToString().Length <= 0) throw new SystemException("【人員判斷】不能為空!");

                            var BarcodeNo = item["BarcodeNo"].ToString();
                            if (BarcodeNo != "")
                            {
                                #region //判斷條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeId, a.BarcodeQty, a.CurrentProdStatus
                                        , b.ItemValue
                                        FROM MES.Barcode a
                                        LEFT JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering'
                                        WHERE a.BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (BarcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                                string ItemValue = "";
                                foreach (var item2 in BarcodeResult)
                                {
                                    if (item2.CurrentProdStatus != "P" && CheckQcMeasureData != "P") throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + item2.ItemValue + ")】目前狀態非良品，不可進行量測上傳!!");
                                    BarcodeId = item2.BarcodeId;
                                    BarcodeQty = item2.BarcodeQty;
                                    ItemValue = item2.ItemValue;
                                    BarcodeIdList.Add(BarcodeId);
                                }
                                #endregion

                                #region //確認條碼是否完工
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.BarcodeProcess a
                                        WHERE a.BarcodeId = @BarcodeId
                                        AND a.FinishDate IS NULL";
                                dynamicParameters.Add("BarcodeId", BarcodeId);

                                var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                if (BarcodeProcessResult.Count() > 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + ItemValue + ")】目前為加工狀態，無法進行量測數據上傳!!");
                                #endregion

                                #region //若為工程檢，檢核條碼是否正在此製程，且已完成加工流程
                                if (SupportProcessFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.CurrentMoProcessId
                                            , b.ProcessAlias
                                            FROM MES.Barcode a
                                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
                                            WHERE a.BarcodeNo = @BarcodeNo";
                                    dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                    var CheckIPQCBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item2 in CheckIPQCBarcodeResult)
                                    {
                                        if (item2.CurrentMoProcessId != MoProcessId) throw new SystemException("條碼【" + BarcodeNo + ")】目前為【" + item2.ProcessAlias + "】完工，無法進行此製程工程檢!!");
                                    }
                                }
                                #endregion

                                #region //確認條碼是否存在品異單且未完成判定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM QMS.AqBarcode a
                                        WHERE a.BarcodeId = @BarcodeId
                                        AND a.JudgeStatus IS NULL";
                                dynamicParameters.Add("BarcodeId", BarcodeId);

                                var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                if (AqBarcodeResult.Count() > 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + ItemValue + ")】目前尚為品異單綁定條碼，且未完成判定!!");
                                #endregion
                            }

                            if (item["SubResponsibleDeptId"].ToString().Length > 0)
                            {
                                #region //判斷副責任單位資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.Department
                                        WHERE DepartmentId = @DepartmentId";
                                dynamicParameters.Add("DepartmentId", Convert.ToInt32(item["SubResponsibleDeptId"]));

                                var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                                if (DepartmentResult.Count() <= 0) throw new SystemException("副責任單位資料錯誤!");
                                #endregion
                            }

                            if (item["SubResponsibleUserId"].ToString().Length > 0)
                            {
                                #region //判斷副責任者資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.[User]
                                        WHERE UserId = @UserId";
                                dynamicParameters.Add("UserId", Convert.ToInt32(item["SubResponsibleUserId"]));

                                var SubResponsibleUserIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (SubResponsibleUserIdResult.Count() <= 0) throw new SystemException("副責任者資料錯誤!");
                                #endregion
                            }

                            #region //INSERT MES.QcBarcode AND UPDATE MES.QcRecord
                            if (BarcodeNo != "")
                            {
                                #region //計算PassQty, NgQty
                                int PassQty = -1;
                                int NgQty = -1;
                                if (item["QcStatus"].ToString() == "P")
                                {
                                    PassQty = BarcodeQty;
                                    NgQty = 0;
                                }
                                else
                                {
                                    PassQty = 0;
                                    NgQty = BarcodeQty;
                                }
                                #endregion

                                #region //INSERT MES.QcBarcode
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcBarcode (QcRecordId, BarcodeProcessId, BarcodeId, PassQty, NgQty, SystemStatus, QcStatus, QcUserId, Remark, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcBarcodeId, INSERTED.BarcodeId
                                        VALUES (@QcRecordId, @BarcodeProcessId, @BarcodeId, @PassQty, @NgQty, @SystemStatus, @QcStatus, @QcUserId, @Remark, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcRecordId,
                                    BarcodeProcessId = (int?)null,
                                    BarcodeId,
                                    PassQty,
                                    NgQty,
                                    SystemStatus = "N",
                                    QcStatus = item["QcStatus"].ToString(),
                                    QcUserId = CreateBy,
                                    Remark = item["Remark"].ToString(),
                                    Status = "Y",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in insertResult)
                                {
                                    QcBarcodeList.Add(item2.QcBarcodeId, item2.BarcodeId);
                                }
                                #endregion

                                #region //UPDATE MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //UPDATE MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        QcStatus = @QcStatus,
                                        QcUserId = @QcUserId,
                                        Remark = @Remark,
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcStatus = item["QcStatus"].ToString(),
                                        QcUserId = CreateBy,
                                        Remark = item["Remark"].ToString(),
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                QcBarcodeList.Add(-1, BarcodeId);
                            }
                            #endregion

                            checkEditFlag = "Y";

                            if (BarcodeNo == "")
                            {
                                if (BarcodeList.IndexOf("Cavity") == -1)
                                {
                                    BarcodeList.Add("Cavity");
                                }
                            }
                            else
                            {
                                if (BarcodeList.IndexOf(BarcodeNo) == -1)
                                {
                                    BarcodeList.Add(BarcodeNo);
                                }
                            }
                        }
                        #endregion

                        if (CheckQcMeasureData != "N" && CheckQcMeasureData != "S" && (barcodeQcRecordJson["barcodeQcRecordInfo"].Count() <= 0 || (barcodeQcRecordJson["barcodeQcRecordInfo"].Count() != BarcodeList.Count())))
                        {
                            throw new SystemException("尚有條碼尚未維護後續異常處理!!");
                        }

                        #region //解析SpreadsheetData
                        var spreadsheetJson = JObject.Parse(QcData);
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {
                            if (item["MeasureValue"] == null) continue;
                            if (item["MeasureValue"].ToString().Length <= 0) throw new SystemException("量測值不能為空!");
                            if (item["QcResult"].ToString() == "F" && (item["NgCode"] == null || item["NgCodeDesc"] == null) && SupportAqFlag == "Y")
                            {
                                if (item["BarcodeNo"] != null)
                                {
                                    throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】之項目【" + item["QcItemName"].ToString() + "】尚未維護異常原因!");
                                }
                                else
                                {
                                    throw new SystemException("項目【" + item["QcItemName"].ToString() + "】尚未維護異常原因!");
                                }
                            }

                            if (item["QcResult"].ToString() == "F") passBool = false;

                            #region //分解10碼
                            if (item["ItemNo"].ToString().Length != 10) throw new SystemException("量測項目【" + item["ItemNo"].ToString() + "】編碼錯誤!!");
                            string MachineNumber = item["ItemNo"].ToString().Substring(3, 3);
                            string QcItemNo = item["ItemNo"].ToString().Substring(0, 3) + item["ItemNo"].ToString().Substring(6, 4);
                            #endregion

                            #region //判斷量測項目資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId
                                    FROM QMS.QcItem a
                                    WHERE a.QcItemNo = @QcItemNo";
                            dynamicParameters.Add("QcItemNo", QcItemNo);

                            var QcItemIdResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcItemIdResult.Count() <= 0) throw new SystemException("量測項目【" + QcItemNo + "】資料錯誤!");

                            int QcItemId = -1;
                            foreach (var item2 in QcItemIdResult)
                            {
                                QcItemId = item2.QcItemId;
                            }
                            #endregion

                            #region //判斷量測機台資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT QmmDetailId
                                    FROM QMS.QmmDetail
                                    WHERE MachineNumber = @MachineNumber";
                            dynamicParameters.Add("MachineNumber", MachineNumber);

                            var QmmDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QmmDetailResult.Count() <= 0) throw new SystemException("量測機台資料錯誤!");

                            int QmmDetailId = -1;
                            foreach (var item2 in QmmDetailResult)
                            {
                                QmmDetailId = item2.QmmDetailId;
                            }
                            #endregion

                            #region //判斷條碼資料是否正確
                            if (item["BarcodeNo"] == null) item["BarcodeNo"] = "";
                            if (item["BarcodeNo"] != null && item["BarcodeNo"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT BarcodeId
                                        FROM MES.Barcode
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (BarcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                                foreach (var item2 in BarcodeResult)
                                {
                                    BarcodeId = item2.BarcodeId;
                                }
                            }
                            else
                            {
                                BarcodeId = null;
                            }
                            #endregion

                            #region //異常原因資料
                            int CauseId = -1;
                            if (item["QcResult"].ToString() == "F" && SupportAqFlag == "Y")
                            {
                                #region //判斷異常原因資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.CauseId
                                        FROM QMS.DefectCause a
                                        WHERE a.CauseNo = @CauseNo";
                                dynamicParameters.Add("CauseNo", item["NgCode"].ToString());

                                var CauseResult = sqlConnection.Query(sql, dynamicParameters);
                                if (CauseResult.Count() <= 0) throw new SystemException("異常原因資料錯誤!");

                                foreach (var item2 in CauseResult)
                                {
                                    CauseId = item2.CauseId;
                                }
                                #endregion

                                #region //維護異常原因項目modal
                                QcNgCode qcNgCode = new QcNgCode
                                {
                                    BarcodeId = BarcodeId,
                                    QcItemId = QcItemId,
                                    CauseId = CauseId,
                                    CauseDesc = item["NgCodeDesc"].ToString()
                                };

                                QcNgCodes.Add(qcNgCode);
                                #endregion

                                //if (item["BarcodeNo"] != null)
                                //{
                                //    #region //更改MES.Barode狀態
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE MES.Barcode SET
                                //            CurrentProdStatus = @CurrentProdStatus,
                                //            LastModifiedDate = @LastModifiedDate,
                                //            LastModifiedBy = @LastModifiedBy
                                //            WHERE BarcodeNo = @BarcodeNo";
                                //    dynamicParameters.AddDynamicParams(
                                //      new
                                //      {
                                //          CurrentProdStatus = "F",
                                //          LastModifiedDate,
                                //          LastModifiedBy,
                                //          BarcodeNo = item["BarcodeNo"].ToString()
                                //      });

                                //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //    #endregion
                                //}
                            }
                            #endregion

                            #region //確認量測人員資料是否正確
                            int QcUserId = -1;
                            if (item["QcUserNo"] != null)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT UserId
                                        FROM BAS.[User]
                                        WHERE UserNo = @UserNo";
                                dynamicParameters.Add("UserNo", item["QcUserNo"].ToString());

                                var QcUserResult = sqlConnection.Query(sql, dynamicParameters);
                                if (QcUserResult.Count() <= 0) throw new SystemException("量測人員資料錯誤!");

                                foreach (var item2 in QcUserResult)
                                {
                                    QcUserId = item2.UserId;
                                }
                            }
                            #endregion

                            #region //確認批號資料是否正確
                            if (item["LotNumber"] != null)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.LotNumber a 
                                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                        WHERE b.MtlItemNo = @MtlItemNo
                                        AND a.LotNumberNo = @LotNumberNo 
                                        AND b.CompanyId = @CompanyId";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("LotNumberNo", item["LotNumber"].ToString());
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                //if (LotNumberResult.Count() <= 0) throw new SystemException("查無批號【" + item["LotNumber"].ToString() + "】相關資料!!");
                            }
                            #endregion

                            if (CheckQcMeasureData == "N")
                            {
                                #region //確認數據資料
                                QcMeasurementData qcMeasurementData = new QcMeasurementData();
                                try
                                {
                                    qcMeasurementData.QcRecordId = QcRecordId;
                                    qcMeasurementData.QcItemId = QcItemId;
                                    qcMeasurementData.QcItemDesc = item["QcItemDesc"] != null ? item["QcItemDesc"].ToString() : null;
                                    qcMeasurementData.DesignValue = item["DesignValue"] != null ? item["DesignValue"].ToString() : null;
                                    qcMeasurementData.UpperTolerance = item["UpperTolerance"] != null ? item["UpperTolerance"].ToString() : null;
                                    qcMeasurementData.LowerTolerance = item["LowerTolerance"] != null ? item["LowerTolerance"].ToString() : null;
                                    qcMeasurementData.ZAxis = item["ZAxis"] != null ? Convert.ToDouble(item["ZAxis"]) : (double?)null;
                                    qcMeasurementData.BarcodeId = BarcodeId;
                                    qcMeasurementData.QmmDetailId = QmmDetailId;
                                    qcMeasurementData.MeasureValue = item["MeasureValue"].ToString();
                                    qcMeasurementData.QcResult = item["QcResult"].ToString();
                                    qcMeasurementData.CauseId = CauseId > 0 ? CauseId : (int?)null;
                                    qcMeasurementData.CauseDesc = item["NgCodeDesc"] != null ? item["NgCodeDesc"].ToString() : "";
                                    qcMeasurementData.Cavity = item["FullCavityNo"] != null ? item["FullCavityNo"].ToString() : null;
                                    qcMeasurementData.MakeCount = Convert.ToInt32(item["MakeCount"]);
                                    qcMeasurementData.CellHeader = item["CellHeader"].ToString();
                                    qcMeasurementData.Row = item["Row"].ToString();
                                    qcMeasurementData.Remark = "";
                                    qcMeasurementData.BallMark = item["BallMark"] != null ? item["BallMark"].ToString() : "";
                                    qcMeasurementData.Unit = item["Unit"] != null ? item["Unit"].ToString() : "";
                                    qcMeasurementData.QcUserId = item["QcUserNo"] != null ? QcUserId : (int?)null;
                                    qcMeasurementData.QcCycleTime = item["QcCycleTime"] != null ? Convert.ToInt32(item["QcCycleTime"]) : (int?)null;
                                    qcMeasurementData.LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : null;
                                    qcMeasurementData.CreateDate = CreateDate;
                                    qcMeasurementData.LastModifiedDate = LastModifiedDate;
                                    qcMeasurementData.CreateBy = CreateBy;
                                    qcMeasurementData.LastModifiedBy = LastModifiedBy;
                                }
                                catch(Exception ex)
                                {
                                    //用於開發人員偵錯
                                }

                                //QcMeasurementData qcMeasurementData = new QcMeasurementData()
                                //{
                                //    QcRecordId = QcRecordId,
                                //    QcItemId = QcItemId,
                                //    QcItemDesc = item["QcItemDesc"] != null ? item["QcItemDesc"].ToString() : null,
                                //    DesignValue = item["DesignValue"] != null ? item["DesignValue"].ToString() : null,
                                //    UpperTolerance = item["UpperTolerance"] != null ? item["UpperTolerance"].ToString() : null,
                                //    LowerTolerance = item["LowerTolerance"] != null ? item["LowerTolerance"].ToString() : null,
                                //    ZAxis = item["ZAxis"] != null ? Convert.ToDouble(item["ZAxis"]) : (double?)null,
                                //    BarcodeId = BarcodeId,
                                //    QmmDetailId = QmmDetailId,
                                //    MeasureValue = item["MeasureValue"].ToString(),
                                //    QcResult = item["QcResult"].ToString(),
                                //    CauseId = CauseId > 0 ? CauseId : (int?)null,
                                //    CauseDesc = item["NgCodeDesc"] != null ? item["NgCodeDesc"].ToString() : "",
                                //    Cavity = item["FullCavityNo"] != null ? item["FullCavityNo"].ToString() : null,
                                //    MakeCount = Convert.ToInt32(item["MakeCount"]),
                                //    CellHeader = item["CellHeader"].ToString(),
                                //    Row = item["Row"].ToString(),
                                //    Remark = "",
                                //    BallMark = item["BallMark"] != null ? item["BallMark"].ToString() : "",
                                //    Unit = item["Unit"] != null ? item["Unit"].ToString() : "",
                                //    QcUserId = item["QcUserNo"] != null ? QcUserId : (int?)null,
                                //    QcCycleTime = item["QcCycleTime"] != null ? Convert.ToInt32(item["QcCycleTime"]) : (int?)null,
                                //    LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : null,
                                //    CreateDate = CreateDate,
                                //    LastModifiedDate = LastModifiedDate,
                                //    CreateBy = CreateBy,
                                //    LastModifiedBy = LastModifiedBy
                                //};
                                #endregion

                                #region //INSERT MES.QcMeasureData
                                try
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO QMS.QcMeasureData (QcRecordId, QcItemId, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance
                                        , ZAxis, BarcodeId, QmmDetailId, MeasureValue, QcResult, CauseId, CauseDesc, Cavity, MakeCount, CellHeader
                                        , Row, QcUserId, QcCycleTime, Remark, BallMark, Unit, LotNumber
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@QcRecordId, @QcItemId, @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance
                                        , @ZAxis, @BarcodeId, @QmmDetailId, @MeasureValue, @QcResult, @CauseId, @CauseDesc, @Cavity, @MakeCount, @CellHeader
                                        , @Row, @QcUserId, @QcCycleTime, @Remark, @BallMark, @Unit, @LotNumber
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                                    var row = sqlConnection.Execute(sql, qcMeasurementData);

                                    rowsAffected += row;
                                }
                                catch (Exception ex)
                                {
                                    throw new SystemException("項目【" + item["ItemNo"].ToString() + "】 條碼【" + item["FullCavityNo"].ToString() +"】新增至QMS.QcMeasureData時錯誤!\n" + ex.Message);
                                }
                                #endregion

                                if (checkEditFlag != "Y")
                                {
                                    checkEditFlag = "P";
                                }
                            }

                            if (item["BarcodeNo"] != null)
                            {
                                if (SpreadSheetBarcodeList.IndexOf(item["BarcodeNo"].ToString()) == -1)
                                {
                                    SpreadSheetBarcodeList.Add(item["BarcodeNo"].ToString());
                                }
                            }
                        }

                        if (SpreadSheetBarcodeList.Count() == 0)
                        {
                            SpreadSheetBarcodeList.Add("Cavity");
                        }

                        if (SpreadSheetBarcodeList.Count() != BarcodeList.Count())
                        {
                            if (checkEditFlag == "Y") checkEditFlag = "P";
                        }

                        if (QcNgCodes.Count() > 0)
                        {
                            QcCauseJsonString = JsonConvert.SerializeObject(QcNgCodes);
                            QcCauseJsonString = "{\"data\":" + QcCauseJsonString + "}";
                        }
                        #endregion

                        #region //處理判定結果
                        List<AqBarcode> AqBarcodes = new List<AqBarcode>();
                        foreach (var item in QcBarcodeList)
                        {
                            foreach (var item2 in barcodeQcRecordJson["barcodeQcRecordInfo"])
                            {
                                #region //取得BARCODE資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT BarcodeNo
                                        FROM MES.Barcode
                                        WHERE BarcodeId = @BarcodeId";
                                dynamicParameters.Add("BarcodeId", item.Value);

                                var barcodeInfoResult = sqlConnection.Query(sql, dynamicParameters);

                                string BarcodeNo = "";
                                foreach (var item3 in barcodeInfoResult)
                                {
                                    BarcodeNo = item3.BarcodeNo;
                                }
                                #endregion

                                if ((item2["BarcodeNo"].ToString() == BarcodeNo || item2["BarcodeNo"] == null) && item2["QcStatus"].ToString() == "F" && SupportAqFlag == "Y")
                                {
                                    #region //建立品異單JSON、UPDATE MES.Barcode
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.QcRecordId, a.MoId, a.MoProcessId, a.CreateBy
                                            , b.DepartmentId ResponsibleDeptId
                                            , c.QcTypeNo
                                            FROM MES.QcRecord a
                                            INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                            INNER JOIN QMS.QcType c ON a.QcTypeId = c.QcTypeId
                                            WHERE a.QcRecordId = @QcRecordId";
                                    dynamicParameters.Add("QcRecordId", QcRecordId);

                                    var result = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item3 in result)
                                    {
                                        AqBarcode aqBarcode = new AqBarcode
                                        {
                                            QcType = item3.QcTypeNo,
                                            MoId = item3.MoId,
                                            QcRecordId = item3.QcRecordId,
                                            QcBarcodeId = item.Key > 0 ? item.Key : (int?)null,
                                            BarcodeId = item.Value,
                                            ConformUserId = item2["ConformUserId"].ToString().Length > 0 ? Convert.ToInt32(item2["ConformUserId"]) : (int?)null,
                                            ResponsibleDeptId = item3.ResponsibleDeptId,
                                            ResponsibleUserId = item3.CreateBy,
                                            SubResponsibleDeptId = item2["SubResponsibleDeptId"].ToString().Length > 0 ? Convert.ToInt32(item2["SubResponsibleDeptId"]) : (int?)null,
                                            SubResponsibleUserId = item2["SubResponsibleUserId"].ToString().Length > 0 ? Convert.ToInt32(item2["SubResponsibleUserId"]) : (int?)null,
                                            ProgrammerUserId = CreateBy,
                                            MoProcessId = null,
                                            DocDate = CreateDate
                                        };

                                        AqBarcodes.Add(aqBarcode);
                                    }
                                    #endregion

                                    if (BarcodeNo.Length > 0)
                                    {
                                        #region //更改MES.Barode狀態
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.Barcode SET
                                                CurrentProdStatus = @CurrentProdStatus,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE BarcodeNo = @BarcodeNo";
                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              CurrentProdStatus = "F",
                                              LastModifiedDate,
                                              LastModifiedBy,
                                              BarcodeNo
                                          });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                        UpdateMoQtyInformation(MoId);
                                        #endregion
                                    }
                                }
                                else if (item2["BarcodeNo"].ToString() == BarcodeNo && item2["QcStatus"].ToString() == "R") //若判定為R(需補正，當站返修)
                                {
                                    #region //取得此次紀錄BarcodeProcessId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.BarcodeProcessId, a.BarcodeId
                                            , b.QcBarcodeId, b.QcStatus
                                            FROM MES.BarcodeProcess a
                                            LEFT JOIN MES.QcBarcode b ON b.BarcodeProcessId = a.BarcodeProcessId
                                            AND b.QcBarcodeId = (
                                              SELECT TOP 1 ba.QcBarcodeId
                                              FROM MES.QcBarcode ba
                                              WHERE ba.BarcodeProcessId = a.BarcodeProcessId
                                              ORDER BY ba.CreateDate DESC
                                            )
                                            INNER JOIN MES.Barcode c ON a.BarcodeId = c.BarcodeId
                                            WHERE a.MoId = @MoId
                                            AND a.MoProcessId = @MoProcessId
                                            AND a.BarcodeId = @BarcodeId
                                            AND a.FinishDate IS NOT NULL
                                            AND c.CurrentProdStatus = 'P'
                                            ORDER BY a.FinishDate DESC";
                                    dynamicParameters.Add("MoId", MoId);
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                    dynamicParameters.Add("BarcodeId", BarcodeId);

                                    var result13 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result13.Count() <= 0) throw new SystemException("製令歷程資料錯誤!");

                                    int BarcodeProcessId = -1;
                                    foreach (var item3 in result13)
                                    {
                                        BarcodeId = item3.BarcodeId;
                                        BarcodeProcessId = item3.BarcodeProcessId;
                                    }
                                    #endregion

                                    #region //更改MES.Barode狀態
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            NextMoProcessId = @NextMoProcessId,
                                            CurrentProdStatus = @CurrentProdStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          NextMoProcessId = MoProcessId,
                                          CurrentProdStatus = "SR",
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          BarcodeNo
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //call 編程API

                                    #endregion

                                    #region //呼叫UpdateMoQtyInformation更新MoProcess Quantity資訊
                                    UpdateMoQtyInformation(MoId);
                                    #endregion
                                }
                                else if (item2["BarcodeNo"].ToString() == BarcodeNo && item2["QcStatus"].ToString() == "M")
                                {
                                    #region //更新MES.Barcode CurrentProdStatus
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            CurrentProdStatus = 'P',
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

                                    #region //將MES.QcBarcode Status UPDATE為N
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcBarcode SET
                                            Status = 'N',
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcBarcodeId = @QcBarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          QcBarcodeId = item.Key
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                        }
                        #endregion

                        if (AqBarcodes.Count() > 0)
                        {
                            #region //開始建立品異單
                            string AbnormalqualityData = JsonConvert.SerializeObject(AqBarcodes);
                            AbnormalqualityData = "{\"data\":" + AbnormalqualityData + "}";
                            string AbnormalProjectList = QcCauseJsonString;

                            dynamicParameters = new DynamicParameters();

                            int AbnormalqualityId = 0;
                            int? QcBarcodeId = 0;
                            int? nullData = null;
                            var ConformUserId = nullData;
                            var SubResponsibleDeptId = nullData;
                            var SubResponsibleUserId = nullData;
                            var ResponsibleSupervisorId = nullData;
                            string DocDate = DateTime.Now.ToString("yyyyMM");

                            var AbnormalqualityJson = JObject.Parse(AbnormalqualityData);
                            var AbnormalProjectListJson = JObject.Parse(AbnormalProjectList);

                            int ResponsibleUserId = -1;
                            int ProgrammerUserId = -1;
                            string QcType = "";
                            #region //品異單單頭建立
                            foreach (var abnorItem in AbnormalqualityJson["data"])
                            {
                                if (Convert.ToInt32(abnorItem["MoId"]) <= 0) throw new SystemException("【製令】不能為空!");
                                if (Convert.ToInt32(abnorItem["ResponsibleDeptId"]) <= 0) throw new SystemException("【責任單位】不能為空!");
                                if (Convert.ToInt32(abnorItem["ResponsibleUserId"]) <= 0) throw new SystemException("【責任者】不能為空!");
                                if (Convert.ToString(abnorItem["QcType"]) != "PVTQC"
                                    && Convert.ToString(abnorItem["QcType"]) != "TQC"
                                    && Convert.ToString(abnorItem["QcType"]) != "MFGQC"
                                    && Convert.ToString(abnorItem["QcType"]) != "CQAQC")
                                {
                                    if (Convert.ToInt32(abnorItem["BarcodeId"]) <= 0) throw new SystemException("【條碼】不能為空!");

                                    if (Convert.ToString(abnorItem["QcType"]) != "NON")
                                    {
                                        if (Convert.ToInt32(abnorItem["QcBarcodeId"]) <= 0) throw new SystemException("【檢驗紀錄編號】不能為空!");
                                    }


                                    #region //判斷條碼是否有進品異
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.AqBarcodeId,a.ProcessStatus,a.JudgeStatus
                                                    FROM QMS.AqBarcode a
								                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(abnorItem["BarcodeId"]));
                                    var resultJudgeStatus = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultJudgeStatus.Count() > 0)
                                    {
                                        string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                        string ProcessStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessStatus;
                                        int AqBarcodeId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).AqBarcodeId;
                                        if (JudgeStatus == "RW")
                                        {
                                            if (ProcessStatus == "I")
                                            {
                                                #region //資料更新
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE QMS.AqBarcode SET
                                                                ProcessStatus = @ProcessStatus,
                                                                LastModifiedBy = @LastModifiedBy,
                                                                LastModifiedDate = @LastModifiedDate
                                                                WHERE AqBarcodeId = @AqBarcodeId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      ProcessStatus = "V",
                                                      ConformUserId,
                                                      LastModifiedBy,
                                                      LastModifiedDate,
                                                      AqBarcodeId
                                                  });
                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion
                                            }
                                        }
                                    }
                                    #endregion
                                }

                                MoId = Convert.ToInt32(abnorItem["MoId"]);
                                QcType = Convert.ToString(abnorItem["QcType"]);
                                ResponsibleUserId = Convert.ToInt32(abnorItem["ResponsibleUserId"]);
                            }

                            #region //單號自動取號
                            string thisDate = CreateDate.ToString("yyyyMM");
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(AbnormalqualityNo))), '000'), 3)) + 1 CurrentNum
                                            FROM QMS.Abnormalquality
								            WHERE AbnormalqualityNo NOT LIKE '%[A-Za-z]%'
                                            AND AbnormalqualityNo LIKE '" + thisDate + "%'";
                            int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                            string AbnormalqualityNo = DocDate + string.Format("{0:000}", currentNum);
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.Abnormalquality (CompanyId, MoId, MoProcessId, AbnormalqualityNo, AbnormalqualityStatus, DocDate, QcType
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.AbnormalqualityId
                                            VALUES (@CompanyId, @MoId, @MoProcessId, @AbnormalqualityNo, @AbnormalqualityStatus, @DocDate, @QcType
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    MoId,
                                    MoProcessId,
                                    AbnormalqualityNo,
                                    AbnormalqualityStatus = "F",
                                    DocDate = DateTime.Now.ToString("yyyyMMdd"),
                                    QcType,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = ResponsibleUserId,
                                    LastModifiedBy = ResponsibleUserId
                                });
                            var insertResult2 = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult2.Count();

                            //取出單頭Id
                            foreach (var AbnormalqualityItem in insertResult2)
                            {
                                AbnormalqualityId = AbnormalqualityItem.AbnormalqualityId;
                            }
                            #endregion

                            #region //品異單單身 - 異常條碼建立
                            foreach (var abnorJsonItem in AbnormalqualityJson["data"])
                            {

                                if (QcType != "PVTQC" && QcType != "TQC" && QcType != "MFGQC" && QcType != "CQAQC")
                                {
                                    if (QcType == "IPQC")
                                    {
                                        #region //判斷條碼是否存在
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                                        FROM MES.Barcode a
                                                        INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                                INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                                        WHERE a.BarcodeId = @BarcodeId
                                                        --AND a.BarcodeStatus = '1' 
                                                        AND　c.FinishDate is not null";
                                        dynamicParameters.Add("BarcodeId", Convert.ToInt32(abnorJsonItem["BarcodeId"]));
                                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(abnorJsonItem["BarcodeId"]) + "】不存在，請重新輸入!");
                                        #endregion
                                    }
                                    else if (QcType == "OQC")
                                    {
                                        #region //判斷條碼是否存在
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                                        FROM MES.Barcode a
                                                        INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                                INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                                        WHERE a.BarcodeId = @BarcodeId
                                                        AND　c.FinishDate is not null";
                                        dynamicParameters.Add("BarcodeId", Convert.ToInt32(abnorJsonItem["BarcodeId"]));
                                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(abnorJsonItem["BarcodeId"]) + "】不存在，請重新輸入!");
                                        #endregion
                                    }


                                    #region //判斷條碼目前是否在品異單
                                    sql = @"SELECT Top 1 a.AqBarcodeId ,a.JudgeStatus 
                                                    FROM QMS.AqBarcode a
                                                    WHERE a.BarcodeId =@BarcodeId
                                                    Order By a.LastModifiedDate DESC
                                                    ";
                                    dynamicParameters.Add("BarcodeId", Convert.ToInt32(abnorJsonItem["BarcodeId"]));
                                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result2.Count() > 0)
                                    {
                                        //判斷品異單判斷結果,如果有值且不是S代表已經判定完成,如果條碼有異常可以開立新的品異單
                                        string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                        if (JudgeStatus == null)
                                        {
                                            throw new SystemException("【條碼Id:" + Convert.ToInt32(abnorJsonItem["BarcodeId"]) + "】目前在品異單判定中，不可以開立品異單");
                                        }
                                        else if (JudgeStatus == "S")
                                        {
                                            throw new SystemException("【條碼Id:" + Convert.ToInt32(abnorJsonItem["BarcodeId"]) + "】目前判定不良品，不可以開立品異單");
                                        }
                                    }
                                    #endregion
                                }

                                #region //判斷資料是否存在
                                #region //資料 - NOT NULL

                                #region //判斷檢驗紀錄編號是否存在
                                if (QcType != "PVTQC" && QcType != "TQC" && QcType != "MFGQC" && QcType != "CQAQC")
                                {
                                    if (QcType != "NON")
                                    {
                                        sql = @"SELECT TOP 1 1
                                                        FROM MES.QcBarcode a
                                                        WHERE a.QcBarcodeId = @QcBarcodeId";
                                        dynamicParameters.Add("QcBarcodeId", Convert.ToInt32(abnorJsonItem["QcBarcodeId"]));
                                        var resultQcBarcodeId = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultQcBarcodeId.Count() <= 0) throw new SystemException("【檢驗紀錄編號:" + abnorJsonItem["QcRecordId"].ToString() + "】不存在，請重新輸入!");
                                    }
                                }
                                #endregion

                                #region //判斷責任單位是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                                FROM BAS.[User] a
                                                INNER JOIN BAS.Department b on a.DepartmentId = b.DepartmentId
                                                WHERE b.DepartmentId = @ResponsibleDeptId";
                                dynamicParameters.Add("ResponsibleDeptId", Convert.ToInt32(abnorJsonItem["ResponsibleDeptId"]));

                                var UserResult = sqlConnection.Query(sql, dynamicParameters);
                                if (UserResult.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                                #endregion

                                #region //判斷責任者是否存在
                                dynamicParameters = new DynamicParameters();

                                sql = @"SELECT TOP 1 1
                                                FROM BAS.[User] a
                                                WHERE a.UserId = @ResponsibleUserId
                                                AND a.Status = 'A'";
                                dynamicParameters.Add("ResponsibleUserId", Convert.ToInt32(abnorJsonItem["ResponsibleUserId"]));

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                                #endregion
                                #endregion

                                #region //資料 - NULL
                                #region //判斷合致對象是否存在
                                if (abnorJsonItem["ConformUserId"].ToString() != "")
                                {
                                    dynamicParameters = new DynamicParameters();

                                    sql = @"SELECT TOP 1 1
                                                    FROM BAS.[User] a
                                                    WHERE a.UserId = @ConformUserId
                                                    AND a.Status = 'A'";
                                    dynamicParameters.Add("ConformUserId", Convert.ToInt32(abnorJsonItem["ConformUserId"]));

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【合致對象】不存在，請重新輸入!");
                                    ConformUserId = Convert.ToInt32(abnorJsonItem["ConformUserId"]);
                                }
                                #endregion

                                #region //判斷副責任單位是否存在
                                if (abnorJsonItem["SubResponsibleDeptId"].ToString() != "")
                                {
                                    dynamicParameters = new DynamicParameters();

                                    sql = @"SELECT TOP 1 1
                                                    FROM BAS.Department a
                                                    WHERE a.DepartmentId = @SubResponsibleDeptId";
                                    dynamicParameters.Add("SubResponsibleDeptId", Convert.ToInt32(abnorJsonItem["SubResponsibleDeptId"]));

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【副責任單位】不存在，請重新輸入!");
                                    SubResponsibleDeptId = Convert.ToInt32(abnorJsonItem["SubResponsibleDeptId"]);
                                }
                                #endregion

                                #region //判斷副責任者是否存在
                                if (abnorJsonItem["SubResponsibleUserId"].ToString() != "")
                                {
                                    dynamicParameters = new DynamicParameters();

                                    sql = @"SELECT TOP 1 1
                                                    FROM BAS.[User] a
                                                    WHERE a.UserId = @SubResponsibleUserId
                                                    AND a.Status = 'A'";
                                    dynamicParameters.Add("SubResponsibleUserId", Convert.ToInt32(abnorJsonItem["SubResponsibleUserId"]));

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【副責任者】不存在，請重新輸入!");

                                    SubResponsibleUserId = Convert.ToInt32(abnorJsonItem["SubResponsibleUserId"]);

                                }
                                #endregion

                                #region //判斷編程者是否存在
                                if (abnorJsonItem["ProgrammerUserId"].ToString() != "")
                                {
                                    dynamicParameters = new DynamicParameters();

                                    sql = @"SELECT TOP 1 1
                                                    FROM BAS.[User] a
                                                    WHERE a.UserId = @ProgrammerUserId
                                                    AND a.Status = 'A'";
                                    dynamicParameters.Add("ProgrammerUserId", Convert.ToInt32(abnorJsonItem["ProgrammerUserId"]));

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【編程者】不存在，請重新輸入!");

                                    ProgrammerUserId = Convert.ToInt32(abnorJsonItem["ProgrammerUserId"]);

                                }
                                #endregion
                                #endregion
                                #endregion

                                if (QcType != "PVTQC" && QcType != "TQC" && QcType != "MFGQC" && QcType != "CQAQC")
                                {
                                    #region //更新異常條碼目前狀態
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                                    CurrentProdStatus = 'F',
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          LastModifiedBy,
                                          LastModifiedDate,
                                          BarcodeId = Convert.ToInt32(abnorJsonItem["BarcodeId"])
                                      });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    BarcodeId = Convert.ToInt32(abnorJsonItem["BarcodeId"]);
                                    if (QcType == "NON")
                                    {
                                        if (abnorJsonItem["QcBarcodeId"].Count() == 0)
                                        {
                                            QcBarcodeId = nullData;
                                        }
                                    }
                                    else
                                    {
                                        QcBarcodeId = Convert.ToInt32(abnorJsonItem["QcBarcodeId"]);
                                    }

                                }
                                else
                                {
                                    BarcodeId = (int?)null;
                                    QcBarcodeId = nullData;
                                }

                                #region //新增 - 品異單單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.AqBarcode (AbnormalqualityId, BarcodeId, QcRecordId, QcBarcodeId, DefectCauseId, DefectCauseDesc, ConformUserId, 
                                        ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, 
                                        RepairCauseId, RepairCauseDesc, RepairCauseUserId,AqBarcodeStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.AqBarcodeId
                                        VALUES (@AbnormalqualityId, @BarcodeId, @QcRecordId, @QcBarcodeId, @DefectCauseId, @DefectCauseDesc, @ConformUserId, 
                                        @ResponsibleDeptId, @ResponsibleUserId, @SubResponsibleDeptId, @SubResponsibleUserId, @ProgrammerUserId, 
                                        @RepairCauseId, @RepairCauseDesc, @RepairCauseUserId, @AqBarcodeStatus, 
                                        @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        AbnormalqualityId,
                                        BarcodeId,
                                        QcRecordId,
                                        QcBarcodeId,
                                        DefectCauseId = nullData,
                                        DefectCauseDesc = nullData,
                                        ConformUserId,
                                        ResponsibleDeptId = Convert.ToInt32(abnorJsonItem["ResponsibleDeptId"]),
                                        ResponsibleUserId = Convert.ToInt32(abnorJsonItem["ResponsibleUserId"]),
                                        SubResponsibleDeptId,
                                        SubResponsibleUserId,
                                        ProgrammerUserId,
                                        RepairCauseId = nullData,
                                        RepairCauseDesc = nullData,
                                        RepairCauseUserId = nullData,
                                        AqBarcodeStatus = 1,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult4 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult4.Count();

                                int AqBarcodeId = 0;
                                foreach (var abnorJsonItem2 in insertResult4)
                                {
                                    AqBarcodeId = abnorJsonItem2.AqBarcodeId;
                                }
                                #endregion

                                foreach (var item4 in AbnormalProjectListJson["data"])
                                {
                                    if (BarcodeId != null && BarcodeId == Convert.ToInt32(item4["BarcodeId"]))
                                    {
                                        if (Convert.ToInt32(item4["CauseId"]) <= 0) throw new SystemException("【不良代碼】不能為空!");
                                        if (item4["CauseDesc"].ToString().Length > 100) throw new SystemException("【不良原因】長度錯誤!");

                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO QMS.AqQcItem (AqBarcodeId, QcItemId, DefectCauseId, DefectCauseDesc, RepairCauseId, RepairCauseDesc
                                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                        OUTPUT INSERTED.AqQcItemId
                                                        VALUES (@AqBarcodeId, @QcItemId, @DefectCauseId, @DefectCauseDesc, @RepairCauseId, @RepairCauseDesc
                                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                AqBarcodeId,
                                                QcItemId = Convert.ToInt32(item4["QcItemId"]),
                                                DefectCauseId = Convert.ToInt32(item4["CauseId"]),
                                                DefectCauseDesc = item4["CauseDesc"].ToString(),
                                                RepairCauseId = nullData,
                                                RepairCauseDesc = nullData,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy = ResponsibleUserId,
                                                LastModifiedBy = ResponsibleUserId
                                            });
                                        var insertAqQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                        rowsAffected += insertAqQcItemResult.Count();
                                    }
                                }

                            }
                            #endregion
                            #endregion
                        }

                        #region //更改單據狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                SpreadsheetData = @SpreadsheetData,
                                QcStatus = @QcStatus,
                                CheckQcMeasureData = @CheckQcMeasureData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SpreadsheetData,
                                QcStatus = passBool == true ? "P" : "F",
                                CheckQcMeasureData = SupportAqFlag == "Y" ? checkEditFlag : "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //若量測上傳完成，透過MAMO通知推播
                        if (checkEditFlag == "Y" || checkEditFlag == "P")
                        {
                            #region //UPDATE MES.QcRecord QcFinishDate
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.QcRecord SET
                                    QcFinishDate = @QcFinishDate,
                                    UrgentFlag = 'N',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcFinishDate = LastModifiedDate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QcRecordId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            if (OriCheckQcMeasureData == "C")
                            {
                                #region //MAMO推播通知
                                string Content = "";

                                #region //取得送測條碼清單
                                string barcodeInfoDesc = "| 條碼 | 刻字 |\n| :-- | :-- |\n";
                                int count = 0;

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeId
                                        , ISNULL(c.ItemValue, '') ItemValue
                                        , d.BarcodeNo
                                        FROM MES.QcReceiptBarcode a 
                                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                        LEFT JOIN MES.BarcodeAttribute c ON a.BarcodeId = c.BarcodeId AND c.MoId = b.MoId AND c.ItemNo = 'Lettering'
                                        INNER JOIN MES.Barcode d ON a.BarcodeId = d.BarcodeId
                                        WHERE a.QcRecordId = @QcRecordId";
                                dynamicParameters.Add("QcRecordId", QcRecordId);

                                var QcReceiptBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in QcReceiptBarcodeResult)
                                {
                                    barcodeInfoDesc += "| " + item.BarcodeNo + " | " + item.ItemValue + " |\n";
                                    count++;
                                }
                                #endregion

                                Content = "### 【量測完成通知】\n" +
                                            "##### 製令: " + WoErpFullNo + "\n" +
                                            "##### 品號: " + MtlItemNo + "\n" +
                                            "##### 品名: " + MtlItemName + "\n" +
                                            "- 量測類型: " + QcTypeNo + QcTypeName + "\n" +
                                            "- 製程: " + ProcessAlias + "\n" +
                                            "- 送檢單位: " + DepartmentNo + " " + DepartmentName + "\n" +
                                            "- 送檢時間: " + ReceiptDate + "\n" +
                                            "- 量測完成時間: " + CreateDate.ToString() + "\n" +
                                            "- 量測單據連結如下:\n" +
                                            "https://bm.zy-tech.com.tw/MesReport/IframeMeasurementRecord?QcRecordId=" + QcRecordId + "\n" +
                                            "- 送檢條碼:\n" + barcodeInfoDesc + "\n";

                                #region //取得標記USER資料(原送測人員部門)
                                List<string> Tags = new List<string>();
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.UserId
                                        , b.UserNo, b.UserName
                                        FROM MAMO.ChannelMembers a 
                                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                        WHERE a.ChannelId = 1
                                        AND b.DepartmentId = @DepartmentId";
                                dynamicParameters.Add("DepartmentId", DepartmentId);

                                var UserResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in UserResult)
                                {
                                    Tags.Add(item.UserNo);
                                }
                                #endregion

                                #region //確認量測檔案內容
                                List<int> Files = new List<int>();

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(a.CurrentFileId, -1) CurrentFileId
                                        , (
                                            SELECT x.FileId
                                            FROM MES.QcRecordFile x 
                                            WHERE x.QcRecordId = a.QcRecordId
                                            FOR JSON PATH, ROOT('data')
                                        ) QcRecordFile
                                        FROM MES.QcRecord a
                                        WHERE a.QcRecordId = @QcRecordId";
                                dynamicParameters.Add("QcRecordId", QcRecordId);

                                var QcFileResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in QcFileResult)
                                {
                                    #region //確認檔案資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BAS.[File]
                                            WHERE FileId = @FileId";
                                    dynamicParameters.Add("FileId", item.CurrentFileId);

                                    var FileResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (FileResult.Count() > 0)
                                    {
                                        Files.Add(item.CurrentFileId);
                                    }
                                    #endregion

                                    if (item.QcRecordFile != null)
                                    {
                                        JObject QcFileJson = JObject.Parse(item.QcRecordFile);
                                        foreach (var item2 in QcFileJson["data"])
                                        {
                                            #region //確認檔案資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 1
                                                    FROM BAS.[File]
                                                    WHERE FileId = @FileId";
                                            dynamicParameters.Add("FileId", Convert.ToInt32(item2["FileId"]));

                                            var FileResult2 = sqlConnection.Query(sql, dynamicParameters);

                                            if (FileResult2.Count() > 0)
                                            {
                                                Files.Add(Convert.ToInt32(item2["FileId"]));
                                            }
                                            #endregion
                                        }
                                    }

                                }
                                #endregion

                                #region //確認推播群組
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ChannelId
                                        FROM MES.QcMamoChannel a
                                        WHERE a.ModeId = @ModeId
                                        AND PushType = 'Y'";
                                dynamicParameters.Add("ModeId", ModeId);

                                var QcMamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                                if (QcMamoChannelResult.Count() <= 0) throw new SystemException("送測群組資料錯誤!!<br>請確認生產模式【" + ModeName + "】是否已設定推播群組!!");

                                int ChannelId = -1;
                                foreach (var item in QcMamoChannelResult)
                                {
                                    ChannelId = item.ChannelId;
                                }
                                #endregion

                                string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId.ToString(), Content, Tags, Files);

                                JObject MamoResultJson = JObject.Parse(MamoResult);
                                if (MamoResultJson["status"].ToString() != "success")
                                {
                                    throw new SystemException(MamoResultJson["msg"].ToString());
                                }
                                #endregion
                            }
                        }
                        #endregion

                        if ((checkEditFlag == "Y" || checkEditFlag == "P") && CurrentFileId > 0 && (QcTypeNo == "PVTQC" || QcTypeNo == "TQC"))
                        {
                            SendQcMail(sqlConnection, QcRecordId, WoErpFullNo, MtlItemNo, MtlItemName, QcTypeNo, QcTypeName
                                , Remark, FileName, ServerPath2, MoId);
                        }

                        if (checkEditFlag == "P" && QcTypeNo == "OQC")
                        {
                            #region //量測完成自動計送信件
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
                            dynamicParameters.Add("SettingSchema", "OQcDataJudge");
                            dynamicParameters.Add("SettingNo", "Y");

                            var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                            if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                            #endregion

                            foreach (var item in resultMailTemplate)
                            {
                                string mailSubject = "【出貨檢】品名:【" + MtlItemName + "】已量測完成，需判定NG品後續流程",
                                    mailContent = HttpUtility.UrlDecode(item.MailContent);

                                #region //Mail內容
                                mailContent = mailContent.Replace("[FinishDate]", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                mailContent = mailContent.Replace("[WoErpFullNo]", WoErpFullNo);
                                mailContent = mailContent.Replace("[MtlItemNo]", MtlItemNo);
                                mailContent = mailContent.Replace("[MtlItemName]", MtlItemName);
                                mailContent = mailContent.Replace("[Remark]", Remark);
                                mailContent = mailContent.Replace("[JudgeHref]", "https://bm.zy-tech.com.tw/QcDataManagement/QcDataManagement?QcRecordId=" + QcRecordId.ToString() + "");
                                mailContent = mailContent.Replace("[QcDataHref]", "https://bm.zy-tech.com.tw/MesReport/IframeMeasurementRecord?QcRecordId=" + QcRecordId.ToString() + "");
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
                                //BaseHelper.MailSend(mailConfig);
                                #endregion
                            }
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

        #region //UploadTempSpreadsheetJson -- 暫存SpreadsheetData -- Ann 2023-04-05
        public string UploadTempSpreadsheetJson(int QcRecordId, string UploadJsonString)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷量測記錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.QcRecord
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!");
                        #endregion

                        #region //確認是否已上傳過量測數據
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMeasureDataResult.Count() > 0) throw new SystemException("此量測記錄單已上傳過量測記錄，無法重複上傳!!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                SpreadsheetData = @SpreadsheetData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SpreadsheetData = UploadJsonString,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
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

        #region //UploadQcGoodsReceipt -- 上傳進貨檢量測數據 -- Ann 2024-04-03
        public string UploadQcGoodsReceipt(int QcRecordId, string QcData, string BarcodeQcRecordData, string SpreadsheetData, string ServerPath2)
        {
            try
            {
                if (QcData.Length <= 0) throw new SystemException("尚未接收到數據!!");
                if (BarcodeQcRecordData.Length <= 0) throw new SystemException("尚未接收到數據!!");
                if (SpreadsheetData.Length <= 0) throw new SystemException("尚未接收完整Spreadsheet數據!!");

                int rowsAffected = 0;
                int? BarcodeId = -1;
                int BarcodeQty = -1;
                List<int?> BarcodeIdList = new List<int?>();
                Dictionary<int, int?> QcBarcodeList = new Dictionary<int, int?>();

                List<QcNgCode> QcNgCodes = new List<QcNgCode>();
                string QcCauseJsonString = "";
                int MoId = -1;
                int QcTypeId = -1;
                int? MoProcessId = -1;
                string ProcessCheckType = "";
                string CheckQcMeasureData = "";
                string checkEditFlag = "N";
                string Remark = "";
                string ngFlag = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否有誤
                        sql = @"SELECT a.CompanyNo
                                FROM BAS.Company a 
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //檢查量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcTypeId, a.MoProcessId, a.CheckQcMeasureData, a.Remark, ISNULL(a.CurrentFileId, -1) CurrentFileId, a.SupportAqFlag
                                , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') ReceiptDate
                                , ISNULL(b.FileName, '') FileName
                                , ISNULL(c.ProcessAlias, '') ProcessAlias
                                , e.DepartmentId
                                , f.DepartmentNo, f.DepartmentName
                                , g.GrDetailId
                                , i.SupplierId
                                FROM MES.QcRecord a
                                LEFT JOIN BAS.[File] b ON a.CurrentFileId = b.FileId
                                LEFT JOIN MES.MoProcess c ON a.MoProcessId = c.MoProcessId
                                INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId
                                INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                                INNER JOIN MES.QcGoodsReceipt g ON a.QcRecordId = g.QcRecordId
                                INNER JOIN SCM.GrDetail h ON g.GrDetailId = h.GrDetailId
                                INNER JOIN SCM.GoodsReceipt i ON h.GrId = i.GrId
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        string FileName = "";
                        int CurrentFileId = -1;
                        string SupportAqFlag = "";
                        string ProcessAlias = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ReceiptDate = "";
                        int DepartmentId = -1;
                        string OriCheckQcMeasureData = "";
                        int GrDetailId = -1;
                        int SupplierId = -1;
                        foreach (var item2 in QcRecordResult)
                        {
                            if (item2.CheckQcMeasureData == "A") throw new SystemException("單據尚未確認收件，無法上傳!!");

                            QcTypeId = item2.QcTypeId;
                            MoProcessId = item2.MoProcessId;
                            CheckQcMeasureData = item2.CheckQcMeasureData == "C" ? "N" : item2.CheckQcMeasureData;
                            Remark = item2.Remark;
                            FileName = item2.FileName;
                            CurrentFileId = item2.CurrentFileId;
                            SupportAqFlag = item2.SupportAqFlag;
                            ProcessAlias = item2.ProcessAlias;
                            DepartmentNo = item2.DepartmentNo;
                            DepartmentName = item2.DepartmentName;
                            ReceiptDate = item2.ReceiptDate.ToString();
                            DepartmentId = item2.DepartmentId;
                            OriCheckQcMeasureData = item2.CheckQcMeasureData;
                            GrDetailId = item2.GrDetailId;
                            SupplierId = item2.SupplierId;
                        }
                        #endregion

                        #region //判斷量測類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                FROM QMS.QcType a 
                                WHERE a.QcTypeId = @QcTypeId";
                        dynamicParameters.Add("QcTypeId", QcTypeId);

                        var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                        string QcTypeNo = "";
                        string QcTypeName = "";
                        string SupportProcessFlag = "";
                        foreach (var item in QcTypeResult)
                        {
                            if (item.SupportProcessFlag == "Y" && MoProcessId <= 0) throw new SystemException("【量測製程】不能為空!");
                            QcTypeNo = item.QcTypeNo;
                            QcTypeName = item.QcTypeName;
                            SupportProcessFlag = item.SupportProcessFlag;
                        }
                        #endregion

                        #region //確認是否已上傳過量測數據
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMeasureDataResult.Count() > 0 && CheckQcMeasureData != "P") throw new SystemException("此量測記錄單已上傳過量測記錄，無法重複上傳!!");
                        #endregion

                        #region //解析BarcodeQcRecordData
                        var barcodeQcRecordJson = JObject.Parse(BarcodeQcRecordData);

                        List<string> BarcodeList = new List<string>();
                        List<string> SpreadSheetBarcodeList = new List<string>();

                        foreach (var item in barcodeQcRecordJson["barcodeQcRecordInfo"])
                        {
                            if (item["QcStatus"].ToString().Length <= 0) throw new SystemException("【人員判斷】不能為空!");

                            #region //Barcode Validation
                            var BarcodeNo = item["BarcodeNo"].ToString();
                            if (BarcodeNo != "")
                            {
                                #region //判斷條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeId, a.BarcodeQty, a.CurrentProdStatus
                                        , b.ItemValue
                                        FROM MES.Barcode a
                                        LEFT JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering'
                                        WHERE a.BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (BarcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                                string ItemValue = "";
                                foreach (var item2 in BarcodeResult)
                                {
                                    if (item2.CurrentProdStatus != "P" && CheckQcMeasureData != "P") throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + item2.ItemValue + ")】目前狀態非良品，不可進行量測上傳!!");
                                    BarcodeId = item2.BarcodeId;
                                    BarcodeQty = item2.BarcodeQty;
                                    ItemValue = item2.ItemValue;
                                    BarcodeIdList.Add(BarcodeId);
                                }
                                #endregion

                                #region //確認條碼是否完工
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.BarcodeProcess a
                                        WHERE a.BarcodeId = @BarcodeId
                                        AND a.FinishDate IS NULL";
                                dynamicParameters.Add("BarcodeId", BarcodeId);

                                var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                if (BarcodeProcessResult.Count() > 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + ItemValue + ")】目前為加工狀態，無法進行量測數據上傳!!");
                                #endregion

                                #region //確認條碼是否存在品異單且未完成判定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM QMS.AqBarcode a
                                        WHERE a.BarcodeId = @BarcodeId
                                        AND a.JudgeStatus IS NULL";
                                dynamicParameters.Add("BarcodeId", BarcodeId);

                                var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                if (AqBarcodeResult.Count() > 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + ItemValue + ")】目前尚為品異單綁定條碼，且未完成判定!!");
                                #endregion
                            }
                            #endregion

                            #region //副責任單位 Validation
                            if (item["SubResponsibleDeptId"].ToString().Length > 0)
                            {
                                #region //判斷副責任單位資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.Department
                                        WHERE DepartmentId = @DepartmentId";
                                dynamicParameters.Add("DepartmentId", Convert.ToInt32(item["SubResponsibleDeptId"]));

                                var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                                if (DepartmentResult.Count() <= 0) throw new SystemException("副責任單位資料錯誤!");
                                #endregion
                            }
                            #endregion

                            #region //副責任USER Validation
                            if (item["SubResponsibleUserId"].ToString().Length > 0)
                            {
                                #region //判斷副責任者資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.[User]
                                        WHERE UserId = @UserId";
                                dynamicParameters.Add("UserId", Convert.ToInt32(item["SubResponsibleUserId"]));

                                var SubResponsibleUserIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (SubResponsibleUserIdResult.Count() <= 0) throw new SystemException("副責任者資料錯誤!");
                                #endregion
                            }
                            #endregion

                            #region //INSERT MES.QcBarcode AND UPDATE MES.QcRecord
                            if (BarcodeNo != "")
                            {
                                #region //計算PassQty, NgQty
                                int PassQty = -1;
                                int NgQty = -1;
                                if (item["QcStatus"].ToString() == "P")
                                {
                                    PassQty = BarcodeQty;
                                    NgQty = 0;
                                }
                                else
                                {
                                    PassQty = 0;
                                    NgQty = BarcodeQty;
                                }
                                #endregion

                                #region //INSERT MES.QcBarcode
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcBarcode (QcRecordId, BarcodeProcessId, BarcodeId, PassQty, NgQty, SystemStatus, QcStatus, QcUserId, Remark, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcBarcodeId, INSERTED.BarcodeId
                                        VALUES (@QcRecordId, @BarcodeProcessId, @BarcodeId, @PassQty, @NgQty, @SystemStatus, @QcStatus, @QcUserId, @Remark, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcRecordId,
                                    BarcodeProcessId = (int?)null,
                                    BarcodeId,
                                    PassQty,
                                    NgQty,
                                    SystemStatus = "N",
                                    QcStatus = item["QcStatus"].ToString(),
                                    QcUserId = CreateBy,
                                    Remark = item["Remark"].ToString(),
                                    Status = "Y",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in insertResult)
                                {
                                    QcBarcodeList.Add(item2.QcBarcodeId, item2.BarcodeId);
                                }
                                #endregion

                                #region //UPDATE MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //UPDATE MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                        QcStatus = @QcStatus,
                                        QcUserId = @QcUserId,
                                        SpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcStatus = item["QcStatus"].ToString(),
                                        QcUserId = CreateBy,
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                QcBarcodeList.Add(-1, BarcodeId);
                            }
                            #endregion

                            checkEditFlag = "Y";

                            if (BarcodeNo == "")
                            {
                                if (BarcodeList.IndexOf("LotNumber") == -1)
                                {
                                    BarcodeList.Add("LotNumber");
                                }
                            }
                            else
                            {
                                if (BarcodeList.IndexOf(BarcodeNo) == -1)
                                {
                                    BarcodeList.Add(BarcodeNo);
                                }
                            }
                        }
                        #endregion

                        if (CheckQcMeasureData != "N" && CheckQcMeasureData != "C" && (barcodeQcRecordJson["barcodeQcRecordInfo"].Count() <= 0 || (barcodeQcRecordJson["barcodeQcRecordInfo"].Count() != BarcodeList.Count())))
                        {
                            throw new SystemException("尚有條碼尚未維護後續異常處理!!");
                        }

                        #region //解析SpreadsheetData
                        var spreadsheetJson = JObject.Parse(QcData);
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {
                            if (item["MeasureValue"] == null) continue;
                            if (item["MeasureValue"].ToString().Length <= 0) throw new SystemException("量測值不能為空!");
                            if (item["QcResult"].ToString() == "F" && (item["NgCode"] == null || item["NgCodeDesc"] == null) && SupportAqFlag == "Y")
                            {
                                if (item["BarcodeNo"] != null)
                                {
                                    throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】之項目【" + item["QcItemName"].ToString() + "】尚未維護異常原因!");
                                }
                                else
                                {
                                    throw new SystemException("項目【" + item["QcItemName"].ToString() + "】尚未維護異常原因!");
                                }
                            }
                            if (item["QcResult"].ToString() == "F")
                            {
                                ngFlag = "NG";
                            }
                            else
                            {
                                if (ngFlag != "NG")
                                {
                                    ngFlag = "PASS";
                                }
                            }

                            #region //分解10碼
                            if (item["ItemNo"].ToString().Length != 10) throw new SystemException("量測項目【" + item["ItemNo"].ToString() + "】編碼錯誤!!");
                            string MachineNumber = item["ItemNo"].ToString().Substring(3, 3);
                            string QcItemNo = item["ItemNo"].ToString().Substring(0, 3) + item["ItemNo"].ToString().Substring(6, 4);
                            #endregion

                            #region //判斷量測項目資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId
                                    FROM QMS.QcItem a
                                    WHERE a.QcItemNo = @QcItemNo";
                            dynamicParameters.Add("QcItemNo", QcItemNo);

                            var QcItemIdResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcItemIdResult.Count() <= 0) throw new SystemException("量測項目【" + QcItemNo + "】資料錯誤!");

                            int QcItemId = -1;
                            foreach (var item2 in QcItemIdResult)
                            {
                                QcItemId = item2.QcItemId;
                            }
                            #endregion

                            #region //判斷量測機台資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT QmmDetailId
                                    FROM QMS.QmmDetail
                                    WHERE MachineNumber = @MachineNumber";
                            dynamicParameters.Add("MachineNumber", MachineNumber);

                            var QmmDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QmmDetailResult.Count() <= 0) throw new SystemException("量測機台資料錯誤!");

                            int QmmDetailId = -1;
                            foreach (var item2 in QmmDetailResult)
                            {
                                QmmDetailId = item2.QmmDetailId;
                            }
                            #endregion

                            #region //判斷條碼資料是否正確
                            if (item["BarcodeNo"] == null) item["BarcodeNo"] = "";
                            if (item["BarcodeNo"] != null && item["BarcodeNo"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT BarcodeId
                                        FROM MES.Barcode
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (BarcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                                foreach (var item2 in BarcodeResult)
                                {
                                    BarcodeId = item2.BarcodeId;
                                }
                            }
                            else
                            {
                                BarcodeId = null;
                            }
                            #endregion

                            #region //異常原因資料
                            int CauseId = -1;
                            if (item["QcResult"].ToString() == "F" && SupportAqFlag == "Y")
                            {
                                #region //判斷異常原因資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.CauseId
                                        FROM QMS.DefectCause a
                                        WHERE a.CauseNo = @CauseNo";
                                dynamicParameters.Add("CauseNo", item["NgCode"].ToString());

                                var CauseResult = sqlConnection.Query(sql, dynamicParameters);
                                if (CauseResult.Count() <= 0) throw new SystemException("異常原因資料錯誤!");

                                foreach (var item2 in CauseResult)
                                {
                                    CauseId = item2.CauseId;
                                }
                                #endregion

                                #region //維護異常原因項目modal
                                QcNgCode qcNgCode = new QcNgCode
                                {
                                    BarcodeId = BarcodeId,
                                    QcItemId = QcItemId,
                                    CauseId = CauseId,
                                    CauseDesc = item["NgCodeDesc"].ToString()
                                };

                                QcNgCodes.Add(qcNgCode);
                                #endregion

                                //if (item["BarcodeNo"] != null)
                                //{
                                //    #region //更改MES.Barode狀態
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE MES.Barcode SET
                                //            CurrentProdStatus = @CurrentProdStatus,
                                //            LastModifiedDate = @LastModifiedDate,
                                //            LastModifiedBy = @LastModifiedBy
                                //            WHERE BarcodeNo = @BarcodeNo";
                                //    dynamicParameters.AddDynamicParams(
                                //      new
                                //      {
                                //          CurrentProdStatus = "F",
                                //          LastModifiedDate,
                                //          LastModifiedBy,
                                //          BarcodeNo = item["BarcodeNo"].ToString()
                                //      });

                                //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //    #endregion
                                //}
                            }
                            #endregion

                            #region //確認量測人員資料是否正確
                            int QcUserId = -1;
                            if (item["QcUserNo"] != null)
                            {
                                if (item["QcUserNo"].ToString() != "null" && item["QcUserNo"].ToString().Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT UserId
                                            FROM BAS.[User]
                                            WHERE UserNo = @UserNo";
                                    dynamicParameters.Add("UserNo", item["QcUserNo"].ToString());

                                    var QcUserResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (QcUserResult.Count() <= 0) throw new SystemException("量測人員資料錯誤!");

                                    foreach (var item2 in QcUserResult)
                                    {
                                        QcUserId = item2.UserId;
                                    }
                                }
                            }
                            #endregion

                            if (CheckQcMeasureData == "N")
                            {
                                #region //INSERT MES.QcMeasureData
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.QcMeasureData (QcRecordId, QcItemId, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, ZAxis, BarcodeId, QmmDetailId, MeasureValue, QcResult, CauseId, CauseDesc, Cavity, LotNumber, MakeCount, CellHeader, Row, QcUserId, QcCycleTime, Remark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@QcRecordId, @QcItemId, @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance, @ZAxis, @BarcodeId, @QmmDetailId, @MeasureValue, @QcResult, @CauseId, @CauseDesc, @Cavity, @LotNumber, @MakeCount, @CellHeader, @Row, @QcUserId, @QcCycleTime, @Remark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcRecordId,
                                        QcItemId,
                                        QcItemDesc = item["QcItemDesc"] != null ? item["QcItemDesc"].ToString() : null,
                                        DesignValue = item["DesignValue"] != null ? item["DesignValue"].ToString() : null,
                                        UpperTolerance = item["UpperTolerance"] != null ? item["UpperTolerance"].ToString() : null,
                                        LowerTolerance = item["LowerTolerance"] != null ? item["LowerTolerance"].ToString() : null,
                                        ZAxis = item["ZAxis"] != null ? item["ZAxis"].ToString() : null,
                                        BarcodeId,
                                        QmmDetailId,
                                        MeasureValue = item["MeasureValue"].ToString(),
                                        QcResult = item["QcResult"].ToString(),
                                        CauseId = CauseId > 0 ? CauseId : (int?)null,
                                        CauseDesc = item["NgCodeDesc"] != null ? item["NgCodeDesc"].ToString() : "",
                                        Cavity = item["FullCavityNo"] != null ? item["FullCavityNo"].ToString() : null,
                                        LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : null,
                                        MakeCount = Convert.ToDouble(item["MakeCount"]),
                                        CellHeader = item["CellHeader"].ToString(),
                                        Row = item["Row"].ToString(),
                                        Remark = "",
                                        QcUserId = item["QcUserNo"] != null ? QcUserId : (int?)null,
                                        QcCycleTime = item["QcCycleTime"] != null ? Convert.ToInt32(item["QcCycleTime"]) : (int?)null,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion

                                if (checkEditFlag != "Y")
                                {
                                    checkEditFlag = "P";
                                }
                            }

                            if (item["BarcodeNo"] != null)
                            {
                                if (SpreadSheetBarcodeList.IndexOf(item["BarcodeNo"].ToString()) == -1)
                                {
                                    SpreadSheetBarcodeList.Add(item["BarcodeNo"].ToString());
                                }
                            }
                        }

                        if (SpreadSheetBarcodeList.Count() == 0)
                        {
                            SpreadSheetBarcodeList.Add("Cavity");
                        }

                        if (SpreadSheetBarcodeList.Count() != BarcodeList.Count())
                        {
                            if (checkEditFlag == "Y") checkEditFlag = "P";
                        }

                        if (QcNgCodes.Count() > 0)
                        {
                            QcCauseJsonString = JsonConvert.SerializeObject(QcNgCodes);
                            QcCauseJsonString = "{\"data\":" + QcCauseJsonString + "}";
                        }
                        #endregion

                        #region //處理判定結果
                        List<AqBarcode> AqBarcodes = new List<AqBarcode>();
                        foreach (var item in QcBarcodeList)
                        {
                            foreach (var item2 in barcodeQcRecordJson["barcodeQcRecordInfo"])
                            {
                                #region //取得BARCODE資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT BarcodeNo
                                        FROM MES.Barcode
                                        WHERE BarcodeId = @BarcodeId";
                                dynamicParameters.Add("BarcodeId", item.Value);

                                var barcodeInfoResult = sqlConnection.Query(sql, dynamicParameters);

                                string BarcodeNo = "";
                                foreach (var item3 in barcodeInfoResult)
                                {
                                    BarcodeNo = item3.BarcodeNo;
                                }
                                #endregion

                                if ((item2["BarcodeNo"].ToString() == BarcodeNo || item2["BarcodeNo"] == null) && item2["QcStatus"].ToString() == "F" && SupportAqFlag == "Y")
                                {
                                    #region //取得不良原因(固定NG001)
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.CauseId, a.CauseDesc
                                            FROM QMS.DefectCause a 
                                            INNER JOIN QMS.DefectClass b ON a.ClassId = b.ClassId
                                            INNER JOIN QMS.DefectGroup c ON b.GroupId = c.GroupId
                                            WHERE c.CompanyId = @CompanyId
                                            AND a.CauseDesc = 'NG001'";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var DefectCauseResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (DefectCauseResult.Count() <= 0) throw new SystemException("異常原因資料有誤!!");

                                    int CauseId = -1;
                                    string CauseDesc = "";
                                    foreach (var item3 in DefectCauseResult)
                                    {
                                        CauseId = item3.CauseId;
                                        CauseDesc = item3.CauseDesc;
                                    }
                                    #endregion

                                    #region //建立品異單JSON、UPDATE MES.Barcode
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.QcRecordId, a.MoId, a.MoProcessId, a.CreateBy
                                            , b.DepartmentId ResponsibleDeptId
                                            , c.QcTypeNo
                                            FROM MES.QcRecord a
                                            INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                            INNER JOIN QMS.QcType c ON a.QcTypeId = c.QcTypeId
                                            WHERE a.QcRecordId = @QcRecordId";
                                    dynamicParameters.Add("QcRecordId", QcRecordId);

                                    var result = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item3 in result)
                                    {
                                        AqBarcode aqBarcode = new AqBarcode
                                        {
                                            QcType = item3.QcTypeNo,
                                            MoId = item3.MoId,
                                            QcRecordId = item3.QcRecordId,
                                            QcBarcodeId = item.Key > 0 ? item.Key : (int?)null,
                                            BarcodeId = item.Value,
                                            ConformUserId = item2["ConformUserId"].ToString().Length > 0 ? Convert.ToInt32(item2["ConformUserId"]) : (int?)null,
                                            ResponsibleDeptId = item3.ResponsibleDeptId,
                                            ResponsibleUserId = item3.CreateBy,
                                            SubResponsibleDeptId = item2["SubResponsibleDeptId"].ToString().Length > 0 ? Convert.ToInt32(item2["SubResponsibleDeptId"]) : (int?)null,
                                            SubResponsibleUserId = item2["SubResponsibleUserId"].ToString().Length > 0 ? Convert.ToInt32(item2["SubResponsibleUserId"]) : (int?)null,
                                            ProgrammerUserId = CreateBy,
                                            MoProcessId = null,
                                            DocDate = CreateDate,
                                            GrDetailId = GrDetailId,
                                            DefectCauseId = CauseId,
                                            DefectCauseDesc = CauseDesc,
                                            SupplierId = SupplierId
                                        };

                                        AqBarcodes.Add(aqBarcode);
                                    }
                                    #endregion

                                    if (BarcodeNo.Length > 0)
                                    {
                                        #region //更改MES.Barode狀態
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.Barcode SET
                                                CurrentProdStatus = @CurrentProdStatus,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE BarcodeNo = @BarcodeNo";
                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              CurrentProdStatus = "F",
                                              LastModifiedDate,
                                              LastModifiedBy,
                                              BarcodeNo
                                          });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                                        UpdateMoQtyInformation(MoId);
                                        #endregion
                                    }
                                }
                                else if (item2["BarcodeNo"].ToString() == BarcodeNo && item2["QcStatus"].ToString() == "R") //若判定為R(需補正，當站返修)
                                {
                                    #region //取得此次紀錄BarcodeProcessId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.BarcodeProcessId, a.BarcodeId
                                            , b.QcBarcodeId, b.QcStatus
                                            FROM MES.BarcodeProcess a
                                            LEFT JOIN MES.QcBarcode b ON b.BarcodeProcessId = a.BarcodeProcessId
                                            AND b.QcBarcodeId = (
                                              SELECT TOP 1 ba.QcBarcodeId
                                              FROM MES.QcBarcode ba
                                              WHERE ba.BarcodeProcessId = a.BarcodeProcessId
                                              ORDER BY ba.CreateDate DESC
                                            )
                                            INNER JOIN MES.Barcode c ON a.BarcodeId = c.BarcodeId
                                            WHERE a.MoId = @MoId
                                            AND a.MoProcessId = @MoProcessId
                                            AND a.BarcodeId = @BarcodeId
                                            AND a.FinishDate IS NOT NULL
                                            AND c.CurrentProdStatus = 'P'
                                            ORDER BY a.FinishDate DESC";
                                    dynamicParameters.Add("MoId", MoId);
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                    dynamicParameters.Add("BarcodeId", BarcodeId);

                                    var result13 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result13.Count() <= 0) throw new SystemException("製令歷程資料錯誤!");

                                    int BarcodeProcessId = -1;
                                    foreach (var item3 in result13)
                                    {
                                        BarcodeId = item3.BarcodeId;
                                        BarcodeProcessId = item3.BarcodeProcessId;
                                    }
                                    #endregion

                                    #region //更改MES.Barode狀態
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            NextMoProcessId = @NextMoProcessId,
                                            CurrentProdStatus = @CurrentProdStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          NextMoProcessId = MoProcessId,
                                          CurrentProdStatus = "SR",
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          BarcodeNo
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //call 編程API

                                    #endregion

                                    #region //呼叫UpdateMoQtyInformation更新MoProcess Quantity資訊
                                    UpdateMoQtyInformation(MoId);
                                    #endregion
                                }
                                else if (item2["BarcodeNo"].ToString() == BarcodeNo && (item2["QcStatus"].ToString() == "P" || item2["QcStatus"].ToString() == "W")) //良品或特採
                                {
                                    if (BarcodeNo.Length > 0)
                                    {
                                        #region //取得目前製程
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.SortNumber CurrentSortNumber
                                                FROM MES.MoProcess a
                                                WHERE a.MoProcessId = @MoProcessId";
                                        dynamicParameters.Add("MoProcessId", MoProcessId);

                                        var result = sqlConnection.Query(sql, dynamicParameters);

                                        int CurrentSortNumber = -1;
                                        foreach (var item3 in result)
                                        {
                                            CurrentSortNumber = item3.CurrentSortNumber;
                                        }
                                        #endregion

                                        #region //取得此製令最大順序製程
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT MAX(a.SortNumber) MaxSortNumber
                                                FROM MES.MoProcess a
                                                WHERE MoId = @MoId";
                                        dynamicParameters.Add("MoId", MoId);

                                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                                        int MaxSortNumber = -1;
                                        foreach (var item3 in result2)
                                        {
                                            MaxSortNumber = item3.MaxSortNumber;
                                        }
                                        #endregion

                                        #region //更新MES.Barcode CurrentProdStatus
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.Barcode SET
                                                CurrentProdStatus = @CurrentProdStatus,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE BarcodeNo = @BarcodeNo";
                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              CurrentProdStatus = "P",
                                              LastModifiedDate,
                                              LastModifiedBy,
                                              BarcodeNo
                                          });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //若為最後一站，依工程檢頻率UPDATE MES.Barcode
                                        if (CurrentSortNumber == MaxSortNumber)
                                        {
                                            if (ProcessCheckType == "1") //全檢
                                            {
                                                //只更新此條碼
                                                #region //UPDATE BARCODE BarcodeStatus、NEXT PROCESS
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE MES.Barcode SET
                                                        BarcodeStatus = @BarcodeStatus,
                                                        NextMoProcessId = @NextMoProcessId,
                                                        LastModifiedDate = @LastModifiedDate,
                                                        LastModifiedBy = @LastModifiedBy
                                                        WHERE BarcodeNo = @BarcodeNo";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      BarcodeStatus = "0",
                                                      NextMoProcessId = -1,
                                                      LastModifiedDate,
                                                      LastModifiedBy,
                                                      BarcodeNo = item.Key
                                                  });

                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion
                                            }
                                            else if (ProcessCheckType == "2") //抽檢
                                            {
                                                //更新整張製令條碼
                                                #region //UPDATE BARCODE BarcodeStatus、NEXT PROCESS
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE MES.Barcode SET
                                                        BarcodeStatus = @BarcodeStatus,
                                                        NextMoProcessId = @NextMoProcessId,
                                                        LastModifiedDate = @LastModifiedDate,
                                                        LastModifiedBy = @LastModifiedBy
                                                        WHERE MoId = @MoId
                                                        AND CurrentMoProcessId = @CurrentMoProcessId";
                                                dynamicParameters.AddDynamicParams(
                                                  new
                                                  {
                                                      BarcodeStatus = "0",
                                                      NextMoProcessId = -1,
                                                      LastModifiedDate,
                                                      LastModifiedBy,
                                                      MoId,
                                                      CurrentMoProcessId = MoProcessId
                                                  });

                                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                                #endregion
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                else if (item2["BarcodeNo"].ToString() == BarcodeNo && item2["QcStatus"].ToString() == "M")
                                {
                                    #region //更新MES.Barcode CurrentProdStatus
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            CurrentProdStatus = 'P',
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

                                    #region //將MES.QcBarcode Status UPDATE為N
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcBarcode SET
                                            Status = 'N',
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcBarcodeId = @QcBarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          QcBarcodeId = item.Key
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                        }
                        #endregion

                        #region //若存在NG品，建立品異單   
                        if (AqBarcodes.Count() > 0)
                        {
                            string AbnormalqualityData = JsonConvert.SerializeObject(AqBarcodes);
                            AbnormalqualityData = "{\"data\":" + AbnormalqualityData + "}";
                            string AbnormalProjectList = QcCauseJsonString;
                            string dataRequest = AddAbnormalqualityPadProject(AbnormalqualityData, AbnormalProjectList, sqlConnection);

                            JObject dataRequestJson = JObject.Parse(dataRequest);
                            if (dataRequestJson["status"].ToString() != "success")
                            {
                                throw new SystemException(dataRequestJson["msg"].ToString());
                            }
                            else
                            {
                                var insertedData = dataRequestJson["data"];
                                int createdAbnormalqualityId = Convert.ToInt32(insertedData[0]["AbnormalqualityId"]);
                                SendIQcAqMail(sqlConnection, QcRecordId, createdAbnormalqualityId, Remark, FileName, ServerPath2);
                            }
                        }
                        #endregion

                        #region //更改單據狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                SpreadsheetData = @SpreadsheetData,
                                CheckQcMeasureData = @CheckQcMeasureData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SpreadsheetData,
                                CheckQcMeasureData = SupportAqFlag == "Y" ? checkEditFlag : "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcGoodsReceipt SET
                                QcStatus = @QcStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcStatus = ngFlag.Length <= 0 ? "A" : ngFlag == "PASS" ? "Y" : "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //若為全良品，更改進貨單身檢驗狀態
                        if (ngFlag == "PASS")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GrDetail SET
                                    QcStatus = '2',
                                    AcceptQty = ReceiptQty,
                                    ReturnQty = 0,
                                    AvailableQty = ReceiptQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    GrDetailId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            checkQcMeasureData = SupportAqFlag == "Y" ? checkEditFlag : "Y"
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

        #region //UploadQcOutsourcingData -- 解析SpreadSheet Data及上傳量測數據  -- GPAI 2025-04-10
        public string UploadQcOutsourcingData(int QcRecordId, string QcData, string BarcodeQcRecordData, string SpreadsheetData)
        {
            try
            {
                if (QcData.Length <= 0) throw new SystemException("尚未接收到數據!!");
                if (BarcodeQcRecordData.Length <= 0) throw new SystemException("尚未接收到數據!!");
                if (SpreadsheetData.Length <= 0) throw new SystemException("尚未接收完整Spreadsheet數據!!");

                int rowsAffected = 0;
                int? BarcodeId = -1;
                int BarcodeQty = -1;
                List<int?> BarcodeIdList = new List<int?>();
                Dictionary<int, int?> QcBarcodeList = new Dictionary<int, int?>();

                List<QcNgCode> QcNgCodes = new List<QcNgCode>();
                string QcCauseJsonString = "";
                int MoId = -1;
                int QcTypeId = -1;
                int? MoProcessId = -1;
                string ProcessCheckType = "";
                string CheckQcMeasureData = "";
                string checkEditFlag = "N";
                string Remark = "";
                string ngFlag = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否有誤
                        sql = @"SELECT a.CompanyNo
                                FROM BAS.Company a 
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //檢查量測紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcTypeId, a.MoProcessId, a.CheckQcMeasureData, a.Remark, ISNULL(a.CurrentFileId, -1) CurrentFileId, a.SupportAqFlag
                                , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') ReceiptDate
                                , ISNULL(b.FileName, '') FileName
                                , ISNULL(c.ProcessAlias, '') ProcessAlias
                                , e.DepartmentId
                                , f.DepartmentNo, f.DepartmentName
                                
                                FROM MES.QcRecord a
                                LEFT JOIN BAS.[File] b ON a.CurrentFileId = b.FileId
                                LEFT JOIN MES.MoProcess c ON a.MoProcessId = c.MoProcessId
                                INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId
                                INNER JOIN BAS.Department f ON e.DepartmentId = f.DepartmentId
                                
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcRecordResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QcRecordResult.Count() <= 0) throw new SystemException("量測記錄資料錯誤!!");

                        string FileName = "";
                        int CurrentFileId = -1;
                        string SupportAqFlag = "";
                        string ProcessAlias = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        string ReceiptDate = "";
                        int DepartmentId = -1;
                        string OriCheckQcMeasureData = "";
                        int GrDetailId = -1;
                        int SupplierId = -1;
                        foreach (var item2 in QcRecordResult)
                        {
                            if (item2.CheckQcMeasureData != "S") throw new SystemException("單據尚未確認是否開始量測，無法上傳!!");

                            QcTypeId = item2.QcTypeId;
                            CheckQcMeasureData = item2.CheckQcMeasureData == "C" ? "N" : item2.CheckQcMeasureData;
                            Remark = item2.Remark;
                            FileName = item2.FileName;
                            CurrentFileId = item2.CurrentFileId;
                            SupportAqFlag = item2.SupportAqFlag;
                            ProcessAlias = item2.ProcessAlias;
                            DepartmentNo = item2.DepartmentNo;
                            DepartmentName = item2.DepartmentName;
                            ReceiptDate = item2.ReceiptDate.ToString();
                            DepartmentId = item2.DepartmentId;
                            OriCheckQcMeasureData = item2.CheckQcMeasureData;
                      
                        }
                        #endregion

                        #region //判斷量測類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ModeId, a.QcTypeNo, a.QcTypeName, a.SupportAqFlag, a.SupportProcessFlag
                                FROM QMS.QcType a 
                                WHERE a.QcTypeId = @QcTypeId";
                        dynamicParameters.Add("QcTypeId", QcTypeId);

                        var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料不能為空!!");

                        string QcTypeNo = "";
                        string QcTypeName = "";
                        string SupportProcessFlag = "";
                        foreach (var item in QcTypeResult)
                        {
                            if (item.SupportProcessFlag == "Y" && MoProcessId <= 0) throw new SystemException("【量測製程】不能為空!");
                            QcTypeNo = item.QcTypeNo;
                            QcTypeName = item.QcTypeName;
                            SupportProcessFlag = item.SupportProcessFlag;
                        }
                        #endregion

                        #region //確認是否已上傳過量測數據
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureData a
                                WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", QcRecordId);

                        var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMeasureDataResult.Count() > 0 && CheckQcMeasureData != "P") throw new SystemException("此量測記錄單已上傳過量測記錄，無法重複上傳!!");
                        #endregion

                        #region //解析BarcodeQcRecordData
                        //var barcodeQcRecordJson = JObject.Parse(BarcodeQcRecordData);

                        //List<string> BarcodeList = new List<string>();
                        //List<string> SpreadSheetBarcodeList = new List<string>();

                        //foreach (var item in barcodeQcRecordJson["barcodeQcRecordInfo"])
                        //{
                        //    if (item["QcStatus"].ToString().Length <= 0) throw new SystemException("【人員判斷】不能為空!");

                        //    #region //Barcode Validation
                        //    var BarcodeNo = item["BarcodeNo"].ToString();
                        //    if (BarcodeNo != "")
                        //    {
                        //        #region //判斷條碼資料是否正確
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT a.BarcodeId, a.BarcodeQty, a.CurrentProdStatus
                        //                , b.ItemValue
                        //                FROM MES.Barcode a
                        //                LEFT JOIN MES.BarcodeAttribute b ON a.BarcodeId = b.BarcodeId AND b.ItemNo = 'Lettering'
                        //                WHERE a.BarcodeNo = @BarcodeNo";
                        //        dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                        //        var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                        //        if (BarcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                        //        string ItemValue = "";
                        //        foreach (var item2 in BarcodeResult)
                        //        {
                        //            if (item2.CurrentProdStatus != "P" && CheckQcMeasureData != "P") throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + item2.ItemValue + ")】目前狀態非良品，不可進行量測上傳!!");
                        //            BarcodeId = item2.BarcodeId;
                        //            BarcodeQty = item2.BarcodeQty;
                        //            ItemValue = item2.ItemValue;
                        //            BarcodeIdList.Add(BarcodeId);
                        //        }
                        //        #endregion

                        //        #region //確認條碼是否完工
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT TOP 1 1
                        //                FROM MES.BarcodeProcess a
                        //                WHERE a.BarcodeId = @BarcodeId
                        //                AND a.FinishDate IS NULL";
                        //        dynamicParameters.Add("BarcodeId", BarcodeId);

                        //        var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                        //        if (BarcodeProcessResult.Count() > 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + ItemValue + ")】目前為加工狀態，無法進行量測數據上傳!!");
                        //        #endregion

                        //        #region //確認條碼是否存在品異單且未完成判定
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT TOP 1 1
                        //                FROM QMS.AqBarcode a
                        //                WHERE a.BarcodeId = @BarcodeId
                        //                AND a.JudgeStatus IS NULL";
                        //        dynamicParameters.Add("BarcodeId", BarcodeId);

                        //        var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        //        if (AqBarcodeResult.Count() > 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "(" + ItemValue + ")】目前尚為品異單綁定條碼，且未完成判定!!");
                        //        #endregion
                        //    }
                        //    #endregion

                        //    #region //副責任單位 Validation
                        //    if (item["SubResponsibleDeptId"].ToString().Length > 0)
                        //    {
                        //        #region //判斷副責任單位資料是否正確
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT TOP 1 1
                        //                FROM BAS.Department
                        //                WHERE DepartmentId = @DepartmentId";
                        //        dynamicParameters.Add("DepartmentId", Convert.ToInt32(item["SubResponsibleDeptId"]));

                        //        var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                        //        if (DepartmentResult.Count() <= 0) throw new SystemException("副責任單位資料錯誤!");
                        //        #endregion
                        //    }
                        //    #endregion

                        //    #region //副責任USER Validation
                        //    if (item["SubResponsibleUserId"].ToString().Length > 0)
                        //    {
                        //        #region //判斷副責任者資料是否正確
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT TOP 1 1
                        //                FROM BAS.[User]
                        //                WHERE UserId = @UserId";
                        //        dynamicParameters.Add("UserId", Convert.ToInt32(item["SubResponsibleUserId"]));

                        //        var SubResponsibleUserIdResult = sqlConnection.Query(sql, dynamicParameters);
                        //        if (SubResponsibleUserIdResult.Count() <= 0) throw new SystemException("副責任者資料錯誤!");
                        //        #endregion
                        //    }
                        //    #endregion

                        //    #region //INSERT MES.QcBarcode AND UPDATE MES.QcRecord
                        //    if (BarcodeNo != "")
                        //    {
                        //        #region //計算PassQty, NgQty
                        //        int PassQty = -1;
                        //        int NgQty = -1;
                        //        if (item["QcStatus"].ToString() == "P")
                        //        {
                        //            PassQty = BarcodeQty;
                        //            NgQty = 0;
                        //        }
                        //        else
                        //        {
                        //            PassQty = 0;
                        //            NgQty = BarcodeQty;
                        //        }
                        //        #endregion

                        //        #region //INSERT MES.QcBarcode
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"INSERT INTO MES.QcBarcode (QcRecordId, BarcodeProcessId, BarcodeId, PassQty, NgQty, SystemStatus, QcStatus, QcUserId, Remark, Status
                        //                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                        //                OUTPUT INSERTED.QcBarcodeId, INSERTED.BarcodeId
                        //                VALUES (@QcRecordId, @BarcodeProcessId, @BarcodeId, @PassQty, @NgQty, @SystemStatus, @QcStatus, @QcUserId, @Remark, @Status
                        //                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        //        dynamicParameters.AddDynamicParams(
                        //        new
                        //        {
                        //            QcRecordId,
                        //            BarcodeProcessId = (int?)null,
                        //            BarcodeId,
                        //            PassQty,
                        //            NgQty,
                        //            SystemStatus = "N",
                        //            QcStatus = item["QcStatus"].ToString(),
                        //            QcUserId = CreateBy,
                        //            Remark = item["Remark"].ToString(),
                        //            Status = "Y",
                        //            CreateDate,
                        //            LastModifiedDate,
                        //            CreateBy,
                        //            LastModifiedBy
                        //        });

                        //        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        //        foreach (var item2 in insertResult)
                        //        {
                        //            QcBarcodeList.Add(item2.QcBarcodeId, item2.BarcodeId);
                        //        }
                        //        #endregion

                        //        #region //UPDATE MES.QcRecord
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"UPDATE MES.QcRecord SET
                        //                SpreadsheetData = @SpreadsheetData,
                        //                LastModifiedDate = @LastModifiedDate,
                        //                LastModifiedBy = @LastModifiedBy
                        //                WHERE QcRecordId = @QcRecordId";
                        //        dynamicParameters.AddDynamicParams(
                        //            new
                        //            {
                        //                SpreadsheetData,
                        //                LastModifiedDate,
                        //                LastModifiedBy,
                        //                QcRecordId
                        //            });

                        //        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //        #endregion
                        //    }
                        //    else
                        //    {
                        //        #region //UPDATE MES.QcRecord
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"UPDATE MES.QcRecord SET
                        //                QcStatus = @QcStatus,
                        //                QcUserId = @QcUserId,
                        //                SpreadsheetData = @SpreadsheetData,
                        //                LastModifiedDate = @LastModifiedDate,
                        //                LastModifiedBy = @LastModifiedBy
                        //                WHERE QcRecordId = @QcRecordId";
                        //        dynamicParameters.AddDynamicParams(
                        //            new
                        //            {
                        //                QcStatus = item["QcStatus"].ToString(),
                        //                QcUserId = CreateBy,
                        //                SpreadsheetData,
                        //                LastModifiedDate,
                        //                LastModifiedBy,
                        //                QcRecordId
                        //            });

                        //        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //        #endregion

                        //        QcBarcodeList.Add(-1, BarcodeId);
                        //    }
                        //    #endregion

                        //    checkEditFlag = "Y";

                        //    if (BarcodeNo == "")
                        //    {
                        //        if (BarcodeList.IndexOf("LotNumber") == -1)
                        //        {
                        //            BarcodeList.Add("LotNumber");
                        //        }
                        //    }
                        //    else
                        //    {
                        //        if (BarcodeList.IndexOf(BarcodeNo) == -1)
                        //        {
                        //            BarcodeList.Add(BarcodeNo);
                        //        }
                        //    }
                        //}
                        #endregion

                        //if (CheckQcMeasureData != "N" && CheckQcMeasureData != "C" && (barcodeQcRecordJson["barcodeQcRecordInfo"].Count() <= 0 || (barcodeQcRecordJson["barcodeQcRecordInfo"].Count() != BarcodeList.Count())))
                        //{
                        //    throw new SystemException("尚有條碼尚未維護後續異常處理!!");
                        //}

                        #region //解析SpreadsheetData
                        var spreadsheetJson = JObject.Parse(QcData);
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {
                            if (item["MeasureValue"] == null) continue;
                            if (item["MeasureValue"].ToString().Length <= 0) throw new SystemException("量測值不能為空!");
                            if (item["QcResult"].ToString() == "F" && (item["NgCode"] == null || item["NgCodeDesc"] == null) && SupportAqFlag == "Y")
                            {
                                if (item["BarcodeNo"] != null)
                                {
                                    throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】之項目【" + item["QcItemName"].ToString() + "】尚未維護異常原因!");
                                }
                                else
                                {
                                    throw new SystemException("項目【" + item["QcItemName"].ToString() + "】尚未維護異常原因!");
                                }
                            }
                            if (item["QcResult"].ToString() == "F")
                            {
                                ngFlag = "NG";
                            }
                            else
                            {
                                if (ngFlag != "NG")
                                {
                                    ngFlag = "PASS";
                                }
                            }

                            #region //分解10碼
                            if (item["ItemNo"].ToString().Length != 10) throw new SystemException("量測項目【" + item["ItemNo"].ToString() + "】編碼錯誤!!");
                            string MachineNumber = item["ItemNo"].ToString().Substring(3, 3);
                            string QcItemNo = item["ItemNo"].ToString().Substring(0, 3) + item["ItemNo"].ToString().Substring(6, 4);
                            #endregion

                            #region //判斷量測項目資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId
                                    FROM QMS.QcItem a
                                    WHERE a.QcItemNo = @QcItemNo";
                            dynamicParameters.Add("QcItemNo", QcItemNo);

                            var QcItemIdResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QcItemIdResult.Count() <= 0) throw new SystemException("量測項目【" + QcItemNo + "】資料錯誤!");

                            int QcItemId = -1;
                            foreach (var item2 in QcItemIdResult)
                            {
                                QcItemId = item2.QcItemId;
                            }
                            #endregion

                            #region //判斷量測機台資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT QmmDetailId
                                    FROM QMS.QmmDetail
                                    WHERE MachineNumber = @MachineNumber";
                            dynamicParameters.Add("MachineNumber", MachineNumber);

                            var QmmDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (QmmDetailResult.Count() <= 0) throw new SystemException("量測機台資料錯誤!");

                            int QmmDetailId = -1;
                            foreach (var item2 in QmmDetailResult)
                            {
                                QmmDetailId = item2.QmmDetailId;
                            }
                            #endregion

                            #region //判斷條碼資料是否正確
                            if (item["BarcodeNo"] == null) item["BarcodeNo"] = "";
                            if (item["BarcodeNo"] != null && item["BarcodeNo"].ToString() != "")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT BarcodeId
                                        FROM MES.Barcode
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (BarcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                                foreach (var item2 in BarcodeResult)
                                {
                                    BarcodeId = item2.BarcodeId;
                                }
                            }
                            else
                            {
                                BarcodeId = null;
                            }
                            #endregion

                            #region //異常原因資料
                            int CauseId = -1;
                            if (item["QcResult"].ToString() == "F" && SupportAqFlag == "Y")
                            {
                                #region //判斷異常原因資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.CauseId
                                        FROM QMS.DefectCause a
                                        WHERE a.CauseNo = @CauseNo";
                                dynamicParameters.Add("CauseNo", item["NgCode"].ToString());

                                var CauseResult = sqlConnection.Query(sql, dynamicParameters);
                                if (CauseResult.Count() <= 0) throw new SystemException("異常原因資料錯誤!");

                                foreach (var item2 in CauseResult)
                                {
                                    CauseId = item2.CauseId;
                                }
                                #endregion

                                #region //維護異常原因項目modal
                                QcNgCode qcNgCode = new QcNgCode
                                {
                                    BarcodeId = BarcodeId,
                                    QcItemId = QcItemId,
                                    CauseId = CauseId,
                                    CauseDesc = item["NgCodeDesc"].ToString()
                                };

                                QcNgCodes.Add(qcNgCode);
                                #endregion

                                //if (item["BarcodeNo"] != null)
                                //{
                                //    #region //更改MES.Barode狀態
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE MES.Barcode SET
                                //            CurrentProdStatus = @CurrentProdStatus,
                                //            LastModifiedDate = @LastModifiedDate,
                                //            LastModifiedBy = @LastModifiedBy
                                //            WHERE BarcodeNo = @BarcodeNo";
                                //    dynamicParameters.AddDynamicParams(
                                //      new
                                //      {
                                //          CurrentProdStatus = "F",
                                //          LastModifiedDate,
                                //          LastModifiedBy,
                                //          BarcodeNo = item["BarcodeNo"].ToString()
                                //      });

                                //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //    #endregion
                                //}
                            }
                            #endregion

                            #region //確認量測人員資料是否正確
                            int QcUserId = -1;
                            if (item["QcUserNo"] != null)
                            {
                                if (item["QcUserNo"].ToString() != "null" && item["QcUserNo"].ToString().Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT UserId
                                            FROM BAS.[User]
                                            WHERE UserNo = @UserNo";
                                    dynamicParameters.Add("UserNo", item["QcUserNo"].ToString());

                                    var QcUserResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (QcUserResult.Count() <= 0) throw new SystemException("量測人員資料錯誤!");

                                    foreach (var item2 in QcUserResult)
                                    {
                                        QcUserId = item2.UserId;
                                    }
                                }
                            }
                            #endregion

                            if (CheckQcMeasureData == "N")
                            {
                                #region //INSERT MES.QcMeasureData
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.QcMeasureData (QcRecordId, QcItemId, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, ZAxis, BarcodeId, QmmDetailId, MeasureValue, QcResult, CauseId, CauseDesc, Cavity, LotNumber, MakeCount, CellHeader, Row, QcUserId, QcCycleTime, Remark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@QcRecordId, @QcItemId, @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance, @ZAxis, @BarcodeId, @QmmDetailId, @MeasureValue, @QcResult, @CauseId, @CauseDesc, @Cavity, @LotNumber, @MakeCount, @CellHeader, @Row, @QcUserId, @QcCycleTime, @Remark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcRecordId,
                                        QcItemId,
                                        QcItemDesc = item["QcItemDesc"] != null ? item["QcItemDesc"].ToString() : null,
                                        DesignValue = item["DesignValue"] != null ? item["DesignValue"].ToString() : null,
                                        UpperTolerance = item["UpperTolerance"] != null ? item["UpperTolerance"].ToString() : null,
                                        LowerTolerance = item["LowerTolerance"] != null ? item["LowerTolerance"].ToString() : null,
                                        ZAxis = item["ZAxis"] != null ? item["ZAxis"].ToString() : null,
                                        BarcodeId,
                                        QmmDetailId,
                                        MeasureValue = item["MeasureValue"].ToString(),
                                        QcResult = item["QcResult"].ToString(),
                                        CauseId = CauseId > 0 ? CauseId : (int?)null,
                                        CauseDesc = item["NgCodeDesc"] != null ? item["NgCodeDesc"].ToString() : "",
                                        Cavity = item["FullCavityNo"] != null ? item["FullCavityNo"].ToString() : null,
                                        LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : null,
                                        MakeCount = Convert.ToDouble(item["MakeCount"]),
                                        CellHeader = item["CellHeader"].ToString(),
                                        Row = item["Row"].ToString(),
                                        Remark = "",
                                        QcUserId = item["QcUserNo"] != null ? QcUserId : (int?)null,
                                        QcCycleTime = item["QcCycleTime"] != null ? Convert.ToInt32(item["QcCycleTime"]) : (int?)null,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion

                                if (checkEditFlag != "Y")
                                {
                                    checkEditFlag = "P";
                                }
                            }

                            //if (item["BarcodeNo"] != null)
                            //{
                            //    if (SpreadSheetBarcodeList.IndexOf(item["BarcodeNo"].ToString()) == -1)
                            //    {
                            //        SpreadSheetBarcodeList.Add(item["BarcodeNo"].ToString());
                            //    }
                            //}
                        }

                        //if (SpreadSheetBarcodeList.Count() == 0)
                        //{
                        //    SpreadSheetBarcodeList.Add("Cavity");
                        //}

                        //if (SpreadSheetBarcodeList.Count() != BarcodeList.Count())
                        //{
                        //    if (checkEditFlag == "Y") checkEditFlag = "P";
                        //}

                        if (QcNgCodes.Count() > 0)
                        {
                            QcCauseJsonString = JsonConvert.SerializeObject(QcNgCodes);
                            QcCauseJsonString = "{\"data\":" + QcCauseJsonString + "}";
                        }
                        #endregion

                        #region //處理判定結果
                        //List<AqBarcode> AqBarcodes = new List<AqBarcode>();
                        //foreach (var item in QcBarcodeList)
                        //{
                        //    foreach (var item2 in barcodeQcRecordJson["barcodeQcRecordInfo"])
                        //    {
                        //        #region //取得BARCODE資訊
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT BarcodeNo
                        //                FROM MES.Barcode
                        //                WHERE BarcodeId = @BarcodeId";
                        //        dynamicParameters.Add("BarcodeId", item.Value);

                        //        var barcodeInfoResult = sqlConnection.Query(sql, dynamicParameters);

                        //        string BarcodeNo = "";
                        //        #endregion

                        //        if (barcodeInfoResult.Count() > 0)
                        //        {
                        //            foreach (var item3 in barcodeInfoResult)
                        //            {
                        //                BarcodeNo = item3.BarcodeNo;
                        //            }


                        //            if ((item2["BarcodeNo"].ToString() == BarcodeNo || item2["BarcodeNo"] == null) && item2["QcStatus"].ToString() == "F" && SupportAqFlag == "Y")
                        //            {
                        //                #region //取得不良原因(固定NG001)
                        //                dynamicParameters = new DynamicParameters();
                        //                sql = @"SELECT a.CauseId, a.CauseDesc
                        //                    FROM QMS.DefectCause a 
                        //                    INNER JOIN QMS.DefectClass b ON a.ClassId = b.ClassId
                        //                    INNER JOIN QMS.DefectGroup c ON b.GroupId = c.GroupId
                        //                    WHERE c.CompanyId = @CompanyId
                        //                    AND a.CauseNo = 'NG001'";
                        //                dynamicParameters.Add("CompanyId", CurrentCompany);

                        //                var DefectCauseResult = sqlConnection.Query(sql, dynamicParameters);

                        //                if (DefectCauseResult.Count() <= 0) throw new SystemException("異常原因資料有誤!!");

                        //                int CauseId = -1;
                        //                string CauseDesc = "";
                        //                foreach (var item3 in DefectCauseResult)
                        //                {
                        //                    CauseId = item3.CauseId;
                        //                    CauseDesc = item3.CauseDesc;
                        //                }
                        //                #endregion

                        //                #region //建立品異單JSON、UPDATE MES.Barcode
                        //                dynamicParameters = new DynamicParameters();
                        //                sql = @"SELECT a.QcRecordId, a.MoId, a.MoProcessId, a.CreateBy
                        //                    , b.DepartmentId ResponsibleDeptId
                        //                    , c.QcTypeNo
                        //                    FROM MES.QcRecord a
                        //                    INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        //                    INNER JOIN QMS.QcType c ON a.QcTypeId = c.QcTypeId
                        //                    WHERE a.QcRecordId = @QcRecordId";
                        //                dynamicParameters.Add("QcRecordId", QcRecordId);

                        //                var result = sqlConnection.Query(sql, dynamicParameters);

                        //                foreach (var item3 in result)
                        //                {
                        //                    AqBarcode aqBarcode = new AqBarcode
                        //                    {
                        //                        QcType = item3.QcTypeNo,
                        //                        MoId = item3.MoId,
                        //                        QcRecordId = item3.QcRecordId,
                        //                        QcBarcodeId = item.Key > 0 ? item.Key : (int?)null,
                        //                        BarcodeId = item.Value,
                        //                        ConformUserId = item2["ConformUserId"].ToString().Length > 0 ? Convert.ToInt32(item2["ConformUserId"]) : (int?)null,
                        //                        ResponsibleDeptId = item3.ResponsibleDeptId,
                        //                        ResponsibleUserId = item3.CreateBy,
                        //                        SubResponsibleDeptId = item2["SubResponsibleDeptId"].ToString().Length > 0 ? Convert.ToInt32(item2["SubResponsibleDeptId"]) : (int?)null,
                        //                        SubResponsibleUserId = item2["SubResponsibleUserId"].ToString().Length > 0 ? Convert.ToInt32(item2["SubResponsibleUserId"]) : (int?)null,
                        //                        ProgrammerUserId = CreateBy,
                        //                        MoProcessId = null,
                        //                        DocDate = CreateDate,
                        //                        GrDetailId = GrDetailId,
                        //                        DefectCauseId = CauseId,
                        //                        DefectCauseDesc = CauseDesc,
                        //                        SupplierId = SupplierId
                        //                    };

                        //                    AqBarcodes.Add(aqBarcode);
                        //                }
                        //                #endregion

                        //                if (BarcodeNo.Length > 0)
                        //                {
                        //                    #region //更改MES.Barode狀態
                        //                    dynamicParameters = new DynamicParameters();
                        //                    sql = @"UPDATE MES.Barcode SET
                        //                        CurrentProdStatus = @CurrentProdStatus,
                        //                        LastModifiedDate = @LastModifiedDate,
                        //                        LastModifiedBy = @LastModifiedBy
                        //                        WHERE BarcodeNo = @BarcodeNo";
                        //                    dynamicParameters.AddDynamicParams(
                        //                      new
                        //                      {
                        //                          CurrentProdStatus = "F",
                        //                          LastModifiedDate,
                        //                          LastModifiedBy,
                        //                          BarcodeNo
                        //                      });

                        //                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                    #endregion

                        //                    #region //呼叫UpdateBarcodeProcessStatus更新MoProcess Quantity資訊
                        //                    UpdateMoQtyInformation(MoId);
                        //                    #endregion
                        //                }
                        //            }
                        //            else if (item2["BarcodeNo"].ToString() == BarcodeNo && item2["QcStatus"].ToString() == "R") //若判定為R(需補正，當站返修)
                        //            {
                        //                #region //取得此次紀錄BarcodeProcessId
                        //                dynamicParameters = new DynamicParameters();
                        //                sql = @"SELECT TOP 1 a.BarcodeProcessId, a.BarcodeId
                        //                    , b.QcBarcodeId, b.QcStatus
                        //                    FROM MES.BarcodeProcess a
                        //                    LEFT JOIN MES.QcBarcode b ON b.BarcodeProcessId = a.BarcodeProcessId
                        //                    AND b.QcBarcodeId = (
                        //                      SELECT TOP 1 ba.QcBarcodeId
                        //                      FROM MES.QcBarcode ba
                        //                      WHERE ba.BarcodeProcessId = a.BarcodeProcessId
                        //                      ORDER BY ba.CreateDate DESC
                        //                    )
                        //                    INNER JOIN MES.Barcode c ON a.BarcodeId = c.BarcodeId
                        //                    WHERE a.MoId = @MoId
                        //                    AND a.MoProcessId = @MoProcessId
                        //                    AND a.BarcodeId = @BarcodeId
                        //                    AND a.FinishDate IS NOT NULL
                        //                    AND c.CurrentProdStatus = 'P'
                        //                    ORDER BY a.FinishDate DESC";
                        //                dynamicParameters.Add("MoId", MoId);
                        //                dynamicParameters.Add("MoProcessId", MoProcessId);
                        //                dynamicParameters.Add("BarcodeId", BarcodeId);

                        //                var result13 = sqlConnection.Query(sql, dynamicParameters);
                        //                if (result13.Count() <= 0) throw new SystemException("製令歷程資料錯誤!");

                        //                int BarcodeProcessId = -1;
                        //                foreach (var item3 in result13)
                        //                {
                        //                    BarcodeId = item3.BarcodeId;
                        //                    BarcodeProcessId = item3.BarcodeProcessId;
                        //                }
                        //                #endregion

                        //                #region //更改MES.Barode狀態
                        //                dynamicParameters = new DynamicParameters();
                        //                sql = @"UPDATE MES.Barcode SET
                        //                    NextMoProcessId = @NextMoProcessId,
                        //                    CurrentProdStatus = @CurrentProdStatus,
                        //                    LastModifiedDate = @LastModifiedDate,
                        //                    LastModifiedBy = @LastModifiedBy
                        //                    WHERE BarcodeNo = @BarcodeNo";
                        //                dynamicParameters.AddDynamicParams(
                        //                  new
                        //                  {
                        //                      NextMoProcessId = MoProcessId,
                        //                      CurrentProdStatus = "SR",
                        //                      LastModifiedDate,
                        //                      LastModifiedBy,
                        //                      BarcodeNo
                        //                  });

                        //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                #endregion

                        //                #region //call 編程API

                        //                #endregion

                        //                #region //呼叫UpdateMoQtyInformation更新MoProcess Quantity資訊
                        //                UpdateMoQtyInformation(MoId);
                        //                #endregion
                        //            }
                        //            else if (item2["BarcodeNo"].ToString() == BarcodeNo && (item2["QcStatus"].ToString() == "P" || item2["QcStatus"].ToString() == "W")) //良品或特採
                        //            {
                        //                if (BarcodeNo.Length > 0)
                        //                {
                        //                    #region //取得目前製程
                        //                    dynamicParameters = new DynamicParameters();
                        //                    sql = @"SELECT a.SortNumber CurrentSortNumber
                        //                        FROM MES.MoProcess a
                        //                        WHERE a.MoProcessId = @MoProcessId";
                        //                    dynamicParameters.Add("MoProcessId", MoProcessId);

                        //                    var result = sqlConnection.Query(sql, dynamicParameters);

                        //                    int CurrentSortNumber = -1;
                        //                    foreach (var item3 in result)
                        //                    {
                        //                        CurrentSortNumber = item3.CurrentSortNumber;
                        //                    }
                        //                    #endregion

                        //                    #region //取得此製令最大順序製程
                        //                    dynamicParameters = new DynamicParameters();
                        //                    sql = @"SELECT MAX(a.SortNumber) MaxSortNumber
                        //                        FROM MES.MoProcess a
                        //                        WHERE MoId = @MoId";
                        //                    dynamicParameters.Add("MoId", MoId);

                        //                    var result2 = sqlConnection.Query(sql, dynamicParameters);

                        //                    int MaxSortNumber = -1;
                        //                    foreach (var item3 in result2)
                        //                    {
                        //                        MaxSortNumber = item3.MaxSortNumber;
                        //                    }
                        //                    #endregion

                        //                    #region //更新MES.Barcode CurrentProdStatus
                        //                    dynamicParameters = new DynamicParameters();
                        //                    sql = @"UPDATE MES.Barcode SET
                        //                        CurrentProdStatus = @CurrentProdStatus,
                        //                        LastModifiedBy = @LastModifiedBy
                        //                        WHERE BarcodeNo = @BarcodeNo";
                        //                    dynamicParameters.AddDynamicParams(
                        //                      new
                        //                      {
                        //                          CurrentProdStatus = "P",
                        //                          LastModifiedDate,
                        //                          LastModifiedBy,
                        //                          BarcodeNo
                        //                      });

                        //                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                    #endregion

                        //                    #region //若為最後一站，依工程檢頻率UPDATE MES.Barcode
                        //                    if (CurrentSortNumber == MaxSortNumber)
                        //                    {
                        //                        if (ProcessCheckType == "1") //全檢
                        //                        {
                        //                            //只更新此條碼
                        //                            #region //UPDATE BARCODE BarcodeStatus、NEXT PROCESS
                        //                            dynamicParameters = new DynamicParameters();
                        //                            sql = @"UPDATE MES.Barcode SET
                        //                                BarcodeStatus = @BarcodeStatus,
                        //                                NextMoProcessId = @NextMoProcessId,
                        //                                LastModifiedDate = @LastModifiedDate,
                        //                                LastModifiedBy = @LastModifiedBy
                        //                                WHERE BarcodeNo = @BarcodeNo";
                        //                            dynamicParameters.AddDynamicParams(
                        //                              new
                        //                              {
                        //                                  BarcodeStatus = "0",
                        //                                  NextMoProcessId = -1,
                        //                                  LastModifiedDate,
                        //                                  LastModifiedBy,
                        //                                  BarcodeNo = item.Key
                        //                              });

                        //                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                            #endregion
                        //                        }
                        //                        else if (ProcessCheckType == "2") //抽檢
                        //                        {
                        //                            //更新整張製令條碼
                        //                            #region //UPDATE BARCODE BarcodeStatus、NEXT PROCESS
                        //                            dynamicParameters = new DynamicParameters();
                        //                            sql = @"UPDATE MES.Barcode SET
                        //                                BarcodeStatus = @BarcodeStatus,
                        //                                NextMoProcessId = @NextMoProcessId,
                        //                                LastModifiedDate = @LastModifiedDate,
                        //                                LastModifiedBy = @LastModifiedBy
                        //                                WHERE MoId = @MoId
                        //                                AND CurrentMoProcessId = @CurrentMoProcessId";
                        //                            dynamicParameters.AddDynamicParams(
                        //                              new
                        //                              {
                        //                                  BarcodeStatus = "0",
                        //                                  NextMoProcessId = -1,
                        //                                  LastModifiedDate,
                        //                                  LastModifiedBy,
                        //                                  MoId,
                        //                                  CurrentMoProcessId = MoProcessId
                        //                              });

                        //                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                            #endregion
                        //                        }
                        //                    }
                        //                    #endregion
                        //                }
                        //            }
                        //            else if (item2["BarcodeNo"].ToString() == BarcodeNo && item2["QcStatus"].ToString() == "M")
                        //            {
                        //                #region //更新MES.Barcode CurrentProdStatus
                        //                dynamicParameters = new DynamicParameters();
                        //                sql = @"UPDATE MES.Barcode SET
                        //                    CurrentProdStatus = 'P',
                        //                    LastModifiedBy = @LastModifiedBy
                        //                    WHERE BarcodeNo = @BarcodeNo";
                        //                dynamicParameters.AddDynamicParams(
                        //                  new
                        //                  {
                        //                      LastModifiedDate,
                        //                      LastModifiedBy,
                        //                      BarcodeNo
                        //                  });

                        //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                #endregion

                        //                #region //將MES.QcBarcode Status UPDATE為N
                        //                dynamicParameters = new DynamicParameters();
                        //                sql = @"UPDATE MES.QcBarcode SET
                        //                    Status = 'N',
                        //                    LastModifiedDate = @LastModifiedDate,
                        //                    LastModifiedBy = @LastModifiedBy
                        //                    WHERE QcBarcodeId = @QcBarcodeId";
                        //                dynamicParameters.AddDynamicParams(
                        //                  new
                        //                  {
                        //                      LastModifiedDate,
                        //                      LastModifiedBy,
                        //                      QcBarcodeId = item.Key
                        //                  });

                        //                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //                #endregion
                        //            }
                        //        }

                        //    }
                        //}
                        #endregion

                        #region //若存在NG品，建立品異單
                        //if (AqBarcodes.Count() > 0)
                        //{
                        //    string AbnormalqualityData = JsonConvert.SerializeObject(AqBarcodes);
                        //    AbnormalqualityData = "{\"data\":" + AbnormalqualityData + "}";
                        //    string AbnormalProjectList = QcCauseJsonString;
                        //    string dataRequest = AddAbnormalqualityPadProject(AbnormalqualityData, AbnormalProjectList, sqlConnection);

                        //    JObject dataRequestJson = JObject.Parse(dataRequest);
                        //    if (dataRequestJson["status"].ToString() != "success")
                        //    {
                        //        throw new SystemException(dataRequestJson["msg"].ToString());
                        //    }
                        //}
                        #endregion

                        #region //更改單據狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.QcRecord SET
                                SpreadsheetData = @SpreadsheetData,
                                CheckQcMeasureData = @CheckQcMeasureData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcRecordId = @QcRecordId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SpreadsheetData,
                                CheckQcMeasureData = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                QcRecordId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //dynamicParameters = new DynamicParameters();
                        //sql = @"UPDATE MES.QcGoodsReceipt SET
                        //        QcStatus = @QcStatus,
                        //        LastModifiedDate = @LastModifiedDate,
                        //        LastModifiedBy = @LastModifiedBy
                        //        WHERE QcRecordId = @QcRecordId";
                        //dynamicParameters.AddDynamicParams(
                        //    new
                        //    {
                        //        QcStatus = ngFlag.Length <= 0 ? "A" : ngFlag == "PASS" ? "Y" : "N",
                        //        LastModifiedDate,
                        //        LastModifiedBy,
                        //        QcRecordId
                        //    });

                        //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //若為全良品，更改進貨單身檢驗狀態
                        //if (ngFlag == "PASS")
                        //{
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"UPDATE SCM.GrDetail SET
                        //            QcStatus = '2',
                        //            AcceptQty = ReceiptQty,
                        //            ReturnQty = 0,
                        //            AvailableQty = ReceiptQty,
                        //            LastModifiedDate = @LastModifiedDate,
                        //            LastModifiedBy = @LastModifiedBy
                        //            WHERE GrDetailId = @GrDetailId";
                        //    dynamicParameters.AddDynamicParams(
                        //        new
                        //        {
                        //            LastModifiedDate,
                        //            LastModifiedBy,
                        //            GrDetailId
                        //        });

                        //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //}
                        #endregion

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

        #region //API
        #region //ReloadQcItemDateForExcel -- 解析Excel更新QcItem -- Ann 2024-04-17
        public string ReloadQcItemDateForExcel(List<QcItem> QcItems, string CompanyNo)
        {
            try
            {
                int rowsAffected = 0;
                int CompanyId = -1;
                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in CompanyResult)
                        {
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //比對EXCEL項目是否存在現有項目，有則INSERT，無則UPDATE
                        foreach (var item in QcItems)
                        {
                            #region //先確認QcClass資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.QcClass
                                    WHERe QcClassId = @QcClassId";
                            dynamicParameters.Add("QcClassId", item.QcClassId);

                            var QcClassResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcClassResult.Count() <= 0) throw new SystemException("品項【" + item.QcItemNo + "】類別【" + item.QcClassId + "】錯誤!!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId
                                    FROM QMS.QcItem a 
                                    INNER JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                                    INNER JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                                    WHERE a.QcClassId = @QcClassId
                                    AND a.QcItemNo = @QcItemNo";
                            dynamicParameters.Add("QcClassId", item.QcClassId);
                            dynamicParameters.Add("QcItemNo", item.QcItemNo);

                            var QcItemResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcItemResult.Count() > 0)
                            {
                                foreach (var item2 in QcItemResult)
                                {
                                    #region //UPDATE QMS.QcItem
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE QMS.QcItem SET
                                            QcItemName = @QcItemName,
                                            QcItemDesc = @QcItemDesc,
                                            QcItemType = @QcItemType,
                                            QcType = @QcType,
                                            Remark = @Remark
                                            WHERE QcItemId = @QcItemId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            item.QcItemName,
                                            item.QcItemDesc,
                                            item.QcItemType,
                                            item.QcType,
                                            item2.QcItemId,
                                            item.Remark
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //INSERT QMS.QcItem
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.QcItem (QcClassId, QcItemNo, QcItemName, QcItemDesc, QcItemType, QcType, Status, Remark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcItemId
                                        VALUES (@QcClassId, @QcItemNo, @QcItemName, @QcItemDesc, @QcItemType, @QcType, @Status, @Remark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.QcClassId,
                                        item.QcItemNo,
                                        item.QcItemName,
                                        item.QcItemDesc,
                                        item.QcItemType,
                                        item.QcType,
                                        item.Status,
                                        item.Remark,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
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

        #region //ReloadQcItemPrincipleDateForExcel -- 解析Excel更新QcItemPrinciple -- Ann 2024-04-17
        public string ReloadQcItemPrincipleDateForExcel(List<QcItemPrinciple> QcItemPrinciples, string CompanyNo)
        {
            try
            {
                int rowsAffected = 0;
                int CompanyId = -1;
                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in CompanyResult)
                        {
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //比對EXCEL項目是否存在現有項目，有則INSERT，無則UPDATE
                        foreach (var item in QcItemPrinciples)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrincipleId
                                    FROM QMS.QcItemPrinciple a 
                                    INNER JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                                    INNER JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                                    WHERE a.QcClassId = @QcClassId
                                    AND a.QmmDetailId = @QmmDetailId
                                    AND a.PrincipleNo = @PrincipleNo";
                            dynamicParameters.Add("QcClassId", item.QcClassId);
                            dynamicParameters.Add("QmmDetailId", item.QmmDetailId);
                            dynamicParameters.Add("PrincipleNo", item.PrincipleNo);

                            var QcItemPrincipleResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcItemPrincipleResult.Count() > 0)
                            {
                                foreach (var item2 in QcItemPrincipleResult)
                                {
                                    #region //UPDATE QMS.QcItemPrinciple
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE QMS.QcItemPrinciple SET
                                            PrincipleDesc = @PrincipleDesc
                                            WHERE PrincipleId = @PrincipleId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            item.PrincipleDesc,
                                            item2.PrincipleId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //INSERT QMS.QcItemPrinciple
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.QcItemPrinciple (QcClassId, QmmDetailId, PrincipleNo, PrincipleDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PrincipleId
                                        VALUES (@QcClassId, @QmmDetailId, @PrincipleNo, @PrincipleDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.QcClassId,
                                        item.QmmDetailId,
                                        item.PrincipleNo,
                                        item.PrincipleDesc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
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

        #region //ReloadPrincipleDetailDateForExcel -- 解析Excel更新PrincipleDetail -- Ann 2024-04-24
        public string ReloadPrincipleDetailDateForExcel(List<PrincipleDetail> PrincipleDetails, string CompanyNo)
        {
            try
            {
                int rowsAffected = 0;
                int CompanyId = -1;
                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in CompanyResult)
                        {
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //比對EXCEL項目是否存在現有項目，有則略過，無則INSERT
                        foreach (var item in PrincipleDetails)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM QMS.PrincipleDetail a 
                                    WHERE a.PrincipleId = @PrincipleId
                                    AND a.PrincipleDesc = @PrincipleDesc";
                            dynamicParameters.Add("PrincipleId", item.PrincipleId);
                            dynamicParameters.Add("PrincipleDesc", item.PrincipleDesc);

                            var PrincipleDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (PrincipleDetailResult.Count() > 0)
                            {
                                continue;
                            }
                            else
                            {
                                #region //INSERT QMS.QcItemPrinciple
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO QMS.PrincipleDetail (PrincipleId, PrincipleDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@PrincipleId, @PrincipleDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.PrincipleId,
                                        item.PrincipleDesc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddAbnormalqualityPadProject -- 建立品異單 
        public string AddAbnormalqualityPadProject(string AbnormalqualityData, string AbnormalProjectList, SqlConnection sqlConnection)
        {
            try
            {
                if (AbnormalqualityData.Length <= 0) throw new SystemException("【品異單資料】不能為空!");

                dynamicParameters = new DynamicParameters();

                int AbnormalqualityId = 0;
                int rowsAffected = 0;
                int? MoId = null;
                int? GrDetailId = null;
                int? MoProcessId = null;
                int? BarcodeId = null;
                int? QcRecordId = null;
                int? QcBarcodeId = null;
                int? nullData = null;
                string QcType = "";
                var ConformUserId = nullData;
                var SubResponsibleDeptId = nullData;
                var SubResponsibleUserId = nullData;
                var ProgrammerUserId = nullData;
                var ResponsibleSupervisorId = nullData;
                string DocDate = DateTime.Now.ToString("yyyyMM");


                var AbnormalqualityJson = JObject.Parse(AbnormalqualityData);

                int ResponsibleUserId = -1;
                #region //品異單單頭建立
                foreach (var item in AbnormalqualityJson["data"])
                {
                    QcType = Convert.ToString(item["QcType"]);
                    ResponsibleUserId = Convert.ToInt32(item["ResponsibleUserId"]);
                    if (QcType == "IPQC" || QcType == "NON")
                    {
                        MoProcessId = Convert.ToInt32(item["MoProcessId"]);
                    }
                    else
                    {
                        MoProcessId = null;
                    }
                    if (Convert.ToInt32(item["ResponsibleDeptId"]) <= 0) throw new SystemException("【責任單位】不能為空!");
                    if (Convert.ToInt32(item["ResponsibleUserId"]) <= 0) throw new SystemException("【責任者】不能為空!");
                    if (QcType != "IQC")
                    {
                        if (Convert.ToInt32(item["MoId"]) <= 0) throw new SystemException("【製令】不能為空!");
                        MoId = Convert.ToInt32(item["MoId"]);
                        if (QcType != "PVTQC")
                        {
                            if (Convert.ToInt32(item["BarcodeId"]) <= 0) throw new SystemException("【條碼】不能為空!");

                            if (QcType != "NON")
                            {
                                if (Convert.ToInt32(item["QcBarcodeId"]) <= 0) throw new SystemException("【檢驗紀錄編號】不能為空!");
                            }


                            #region //判斷條碼是否有進品異
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.AqBarcodeId,a.ProcessStatus,a.JudgeStatus
                                        FROM QMS.AqBarcode a
								        WHERE BarcodeId = @BarcodeId";
                            dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                            var resultJudgeStatus = sqlConnection.Query(sql, dynamicParameters);
                            if (resultJudgeStatus.Count() > 0)
                            {
                                string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                                string ProcessStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ProcessStatus;
                                int AqBarcodeId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).AqBarcodeId;
                                if (JudgeStatus == "RW")
                                {
                                    if (ProcessStatus == "I")
                                    {
                                        #region //資料更新
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE QMS.AqBarcode SET
                                                    ProcessStatus = @ProcessStatus,
                                                    LastModifiedBy = @LastModifiedBy,
                                                    LastModifiedDate = @LastModifiedDate
                                                    WHERE AqBarcodeId = @AqBarcodeId";
                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              ProcessStatus = "V",
                                              ConformUserId,
                                              LastModifiedBy,
                                              LastModifiedDate,
                                              AqBarcodeId
                                          });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                    else
                    {
                        if (Convert.ToInt32(item["GrDetailId"]) <= 0) throw new SystemException("【銷貨單】不能為空!");
                        GrDetailId = Convert.ToInt32(item["GrDetailId"]);
                        #region 判斷進貨單是否可以開立
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.AqBarcodeStatus
                                        FROM QMS.Abnormalquality a
                                        INNER JOIN QMS.AqBarcode b on a.AbnormalqualityId = b.AbnormalqualityId
                                        WHERE a.GrDetailId = @GrDetailId
                                        AND b.JudgeConfirm != 'Y'";
                        dynamicParameters.Add("GrDetailId", GrDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該進貨單目前尚存品異單據未完成判定!!不能開立新品異單");
                        #endregion

                    }
                }

                #region //單號自動取號
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(AbnormalqualityNo))), '000'), 3)) + 1 CurrentNum
                                FROM QMS.Abnormalquality
								WHERE AbnormalqualityNo NOT LIKE '%[A-Za-z]%'";
                int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                string AbnormalqualityNo = DocDate + string.Format("{0:000}", currentNum);
                #endregion

                dynamicParameters = new DynamicParameters();
                sql = @"INSERT INTO QMS.Abnormalquality (CompanyId, MoId, GrDetailId, MoProcessId, AbnormalqualityNo, AbnormalqualityStatus, DocDate, QcType
                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                        OUTPUT INSERTED.AbnormalqualityId
                        VALUES (@CompanyId, @MoId, @GrDetailId, @MoProcessId, @AbnormalqualityNo, @AbnormalqualityStatus, @DocDate, @QcType
                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        CompanyId = CurrentCompany,
                        MoId,
                        GrDetailId,
                        MoProcessId,
                        AbnormalqualityNo,
                        AbnormalqualityStatus = "F",
                        DocDate = DateTime.Now.ToString("yyyyMMdd"),
                        QcType,
                        CreateDate,
                        LastModifiedDate,
                        CreateBy = ResponsibleUserId,
                        LastModifiedBy = ResponsibleUserId
                    });
                var insertResult = sqlConnection.Query(sql, dynamicParameters);
                rowsAffected += insertResult.Count();

                //取出單頭Id
                foreach (var item in insertResult)
                {
                    AbnormalqualityId = item.AbnormalqualityId;
                }
                #endregion

                #region //品異單單身 - 異常條碼建立
                foreach (var item in AbnormalqualityJson["data"])
                {

                    if (QcType != "PVTQC" && QcType != "IQC")
                    {

                        if (QcType == "IPQC")
                        {
                            #region //判斷條碼是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                            FROM MES.Barcode a
                                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                    INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                            WHERE a.BarcodeId = @BarcodeId
                                            --AND a.BarcodeStatus = '1' 
                                            AND　c.FinishDate is not null";
                            dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】不存在，請重新輸入!");
                            #endregion
                        }
                        else if (QcType == "OQC")
                        {
                            #region //判斷條碼是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeNo,a.BarcodeStatus
                                            FROM MES.Barcode a
                                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
		                                    INNER JOIN MES.BarcodeProcess c ON b.MoProcessId = c.MoProcessId and a.BarcodeId =c.BarcodeId
                                            WHERE a.BarcodeId = @BarcodeId
                                            AND　c.FinishDate is not null";
                            dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】不存在，請重新輸入!");
                            #endregion
                        }


                        #region //判斷條碼目前是否在品異單
                        sql = @"SELECT Top 1 a.AqBarcodeId ,a.JudgeStatus 
                                        FROM QMS.AqBarcode a
                                        WHERE a.BarcodeId =@BarcodeId
                                        Order By a.LastModifiedDate DESC
                                        ";
                        dynamicParameters.Add("BarcodeId", Convert.ToInt32(item["BarcodeId"]));
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {
                            //判斷品異單判斷結果,如果有值且不是S代表已經判定完成,如果條碼有異常可以開立新的品異單
                            string JudgeStatus = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JudgeStatus;
                            if (JudgeStatus == null)
                            {
                                throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】目前在品異單判定中，不可以開立品異單");
                            }
                            else if (JudgeStatus == "S")
                            {
                                throw new SystemException("【條碼Id:" + Convert.ToInt32(item["BarcodeId"]) + "】目前判定不良品，不可以開立品異單");
                            }
                        }
                        #endregion
                    }

                    #region //判斷資料是否存在
                    #region //資料 - NOT NULL

                    #region //判斷檢驗紀錄編號是否存在
                    if (QcType != "PVTQC" && QcType != "IQC")
                    {
                        if (QcType != "NON")
                        {
                            sql = @"SELECT TOP 1 1
                                        FROM MES.QcBarcode a
                                        WHERE a.QcBarcodeId = @QcBarcodeId";
                            dynamicParameters.Add("QcBarcodeId", Convert.ToInt32(item["QcBarcodeId"]));
                            var resultQcBarcodeId = sqlConnection.Query(sql, dynamicParameters);
                            if (resultQcBarcodeId.Count() <= 0) throw new SystemException("【檢驗紀錄編號:" + item["QcRecordId"].ToString() + "】不存在，請重新輸入!");
                        }
                    }
                    #endregion

                    #region //判斷責任單位是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b on a.DepartmentId = b.DepartmentId
                                WHERE b.DepartmentId = @ResponsibleDeptId";
                    dynamicParameters.Add("ResponsibleDeptId", Convert.ToInt32(item["ResponsibleDeptId"]));

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【責任單位】不存在，請重新輸入!");
                    #endregion

                    #region //判斷責任者是否存在
                    dynamicParameters = new DynamicParameters();

                    sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ResponsibleUserId
                                AND a.Status = 'A'";
                    dynamicParameters.Add("ResponsibleUserId", Convert.ToInt32(item["ResponsibleUserId"]));

                    result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【責任者】不存在，請重新輸入!");
                    #endregion
                    #endregion

                    #region //資料 - NULL
                    #region //判斷合致對象是否存在
                    if (item["ConformUserId"].ToString() != "")
                    {
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ConformUserId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("ConformUserId", Convert.ToInt32(item["ConformUserId"]));

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【合致對象】不存在，請重新輸入!");
                        ConformUserId = Convert.ToInt32(item["ConformUserId"]);
                    }
                    #endregion

                    #region //判斷副責任單位是否存在
                    if (item["SubResponsibleDeptId"].ToString() != "")
                    {
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @SubResponsibleDeptId";
                        dynamicParameters.Add("SubResponsibleDeptId", Convert.ToInt32(item["SubResponsibleDeptId"]));

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【副責任單位】不存在，請重新輸入!");
                        SubResponsibleDeptId = Convert.ToInt32(item["SubResponsibleDeptId"]);
                    }
                    #endregion

                    #region //判斷副責任者是否存在
                    if (item["SubResponsibleUserId"].ToString() != "")
                    {
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @SubResponsibleUserId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("SubResponsibleUserId", Convert.ToInt32(item["SubResponsibleUserId"]));

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【副責任者】不存在，請重新輸入!");

                        SubResponsibleUserId = Convert.ToInt32(item["SubResponsibleUserId"]);

                    }
                    #endregion

                    #region //判斷編程者是否存在
                    if (item["ProgrammerUserId"].ToString() != "")
                    {
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User] a
                                WHERE a.UserId = @ProgrammerUserId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("ProgrammerUserId", Convert.ToInt32(item["ProgrammerUserId"]));

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【編程者】不存在，請重新輸入!");

                        ProgrammerUserId = Convert.ToInt32(item["ProgrammerUserId"]);

                    }
                    #endregion
                    #endregion
                    #endregion



                    if (QcType == "PVTQC")
                    {
                        BarcodeId = nullData;
                        QcBarcodeId = nullData;



                    }
                    else if (QcType == "IQC")
                    {
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 1 
                                        FROM MES.QcRecord a
                                        WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", Convert.ToInt32(item["QcRecordId"]));

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【檢驗紀錄編號】不存在，請重新輸入!");
                        QcRecordId = Convert.ToInt32(item["QcRecordId"]);
                    }
                    else
                    {
                        #region //更新異常條碼目前狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Barcode SET
                                        CurrentProdStatus = 'F',
                                        LastModifiedBy = @LastModifiedBy,
                                        LastModifiedDate = @LastModifiedDate
                                        WHERE BarcodeId = @BarcodeId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              LastModifiedBy,
                              LastModifiedDate,
                              BarcodeId = Convert.ToInt32(item["BarcodeId"])
                          });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        BarcodeId = Convert.ToInt32(item["BarcodeId"]);
                        if (QcType == "NON")
                        {
                            if (item["QcBarcodeId"].Count() == 0)
                            {
                                QcBarcodeId = nullData;
                            }
                        }
                        else
                        {
                            QcBarcodeId = Convert.ToInt32(item["QcBarcodeId"]);
                        }
                    }

                    #region //新增 - 品異單單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"INSERT INTO QMS.AqBarcode (AbnormalqualityId, BarcodeId, QcRecordId, QcBarcodeId, DefectCauseId, DefectCauseDesc, ConformUserId, 
                            ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, 
                            RepairCauseId, RepairCauseDesc, RepairCauseUserId,AqBarcodeStatus, SupplierId
                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                            OUTPUT INSERTED.AqBarcodeId
                            VALUES (@AbnormalqualityId, @BarcodeId, @QcRecordId, @QcBarcodeId, @DefectCauseId, @DefectCauseDesc, @ConformUserId, 
                            @ResponsibleDeptId, @ResponsibleUserId, @SubResponsibleDeptId, @SubResponsibleUserId, @ProgrammerUserId, 
                            @RepairCauseId, @RepairCauseDesc, @RepairCauseUserId, @AqBarcodeStatus, @SupplierId,
                            @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                    dynamicParameters.AddDynamicParams(
                        new
                        {
                            AbnormalqualityId,
                            BarcodeId,
                            QcRecordId,
                            QcBarcodeId,
                            DefectCauseId = QcType != "IQC" ? nullData : item["DefectCauseId"].ToString().Length > 0 ? Convert.ToInt32(item["DefectCauseId"]) : nullData,
                            DefectCauseDesc = QcType != "IQC" ? (string)null : item["DefectCauseDesc"].ToString().Length > 0 ? item["DefectCauseDesc"].ToString() : (string)null,
                            ConformUserId,
                            ResponsibleDeptId = Convert.ToInt32(item["ResponsibleDeptId"]),
                            ResponsibleUserId = Convert.ToInt32(item["ResponsibleUserId"]),
                            SubResponsibleDeptId,
                            SubResponsibleUserId,
                            ProgrammerUserId,
                            RepairCauseId = nullData,
                            RepairCauseDesc = nullData,
                            RepairCauseUserId = nullData,
                            AqBarcodeStatus = 1,
                            SupplierId = item["SupplierId"].ToString().Length > 0 ? Convert.ToInt32(item["SupplierId"]) : (int?)null,
                            CreateDate,
                            LastModifiedDate,
                            CreateBy,
                            LastModifiedBy
                        });
                    var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                    rowsAffected += insertResult2.Count();

                    int AqBarcodeId = 0;
                    foreach (var item2 in insertResult2)
                    {
                        AqBarcodeId = item2.AqBarcodeId;
                    }
                    #endregion

                    if (QcType != "IQC")
                    {
                        if (AbnormalProjectList.Length <= 0) throw new SystemException("【量測異常資料】不能為空!");
                        var AbnormalProjectListJson = JObject.Parse(AbnormalProjectList);

                        foreach (var item2 in AbnormalProjectListJson["data"])
                        {
                            if (Convert.ToInt32(item2["CauseId"]) <= 0) throw new SystemException("【不良代碼】不能為空!");
                            if (item2["CauseDesc"].ToString().Length > 100) throw new SystemException("【不良原因】長度錯誤!");

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.AqQcItem (AqBarcodeId, QcItemId, DefectCauseId, DefectCauseDesc, RepairCauseId, RepairCauseDesc
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.AqQcItemId
                                    VALUES (@AqBarcodeId, @QcItemId, @DefectCauseId, @DefectCauseDesc, @RepairCauseId, @RepairCauseDesc
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    AqBarcodeId,
                                    QcItemId = Convert.ToInt32(item2["QcItemId"]),
                                    DefectCauseId = Convert.ToInt32(item2["CauseId"]),
                                    DefectCauseDesc = item2["CauseDesc"].ToString(),
                                    RepairCauseId = nullData,
                                    RepairCauseDesc = nullData,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = ResponsibleUserId,
                                    LastModifiedBy = ResponsibleUserId
                                });
                            var insertAqQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertAqQcItemResult.Count();
                        }
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

        #region //GetAutomationDataForMTF -- 取得自動化MTF資料表數據 -- WuTC 2024-05-06
        public string GetAutomationDataForMTF(string Company = "", string AutomationRecordId = "")
        {

            try
            {
                bool bInput = true;
                string[] spID = AutomationRecordId.Split(',');

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", Company);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            CurrentCompany = item.CompanyId;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(AutomationConnectionStrings))
                        {
                            #region//取得自動化資料庫中的資料，依拋過來的RecordId查詢
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                    FROM MTF_RawData_TEST
                                    WHERE ID IN (" + spID[0];
                            for (int i = 1; i < spID.Length; i++)
                            {
                                sql += "," + spID[i];
                            }
                            sql += ")";
                            var AutomationResult = sqlConnection2.Query(sql);
                            if (AutomationResult.Count() == 0) throw new SystemException("【自動化】查無IOT資訊!");
                            #endregion

                            #region //宣告自動化MTF資料庫欄位
                            string MoProcessNo = "";
                            string MachineId = "";
                            string TrayNo = "";
                            string UserNo = "";
                            string LocationX = "";
                            string LocationY = "";
                            string MfgOrder = "";
                            string Class = "";
                            string Date_Time = "";
                            string Sup_LotCode = "";
                            #endregion

                            foreach (var DataRow in AutomationResult)
                            {
                                List<string> measureValues = new List<string>();
                                List<string> columnName = new List<string>();
                                #region //取得自動化MTF資料庫欄位資料
                                MachineId = DataRow.MachineSN ?? null;
                                UserNo = DataRow.Operater;
                                MfgOrder = DataRow.MfgOrder;
                                MoProcessNo = DataRow.Process ?? null;
                                TrayNo = DataRow.MTF_TraySN;
                                LocationX = DataRow.Location_X;
                                LocationY = DataRow.Location_Y;
                                Sup_LotCode = DataRow.Sup_LotCode;
                                Class = DataRow.Class;

                                #region //measureValue.add
                                measureValues.Add(DataRow.CheckFreq);
                                measureValues.Add(DataRow.FBL);
                                measureValues.Add(DataRow.DOF_Index);
                                measureValues.Add(DataRow.DOF_minus);
                                measureValues.Add(DataRow.DOF_plus);
                                measureValues.Add(DataRow.DOF_T);
                                measureValues.Add(DataRow.GDOF_Index);
                                measureValues.Add(DataRow.GDOF_minus);
                                measureValues.Add(DataRow.GDOF_plus);
                                measureValues.Add(DataRow.GDOF_T);
                                measureValues.Add(DataRow.EFL);
                                measureValues.Add(DataRow.W_1S);
                                measureValues.Add(DataRow.W_1T);
                                measureValues.Add(DataRow.W_2S);
                                measureValues.Add(DataRow.W_2T);
                                measureValues.Add(DataRow.W_3S);
                                measureValues.Add(DataRow.W_3T);
                                measureValues.Add(DataRow.W_4S);
                                measureValues.Add(DataRow.W_4T);
                                measureValues.Add(DataRow.W_5S);
                                measureValues.Add(DataRow.W_5T);
                                measureValues.Add(DataRow.W_6S);
                                measureValues.Add(DataRow.W_6T);
                                measureValues.Add(DataRow.W_7S);
                                measureValues.Add(DataRow.W_7T);
                                measureValues.Add(DataRow.W_8S);
                                measureValues.Add(DataRow.W_8T);
                                measureValues.Add(DataRow.W_9S);
                                measureValues.Add(DataRow.W_9T);
                                measureValues.Add(DataRow.W_10S);
                                measureValues.Add(DataRow.W_10T);
                                measureValues.Add(DataRow.W_11S);
                                measureValues.Add(DataRow.W_11T);
                                measureValues.Add(DataRow.W_12S);
                                measureValues.Add(DataRow.W_12T);
                                measureValues.Add(DataRow.W_13S);
                                measureValues.Add(DataRow.W_13T);
                                measureValues.Add(DataRow.W_14S);
                                measureValues.Add(DataRow.W_14T);
                                measureValues.Add(DataRow.W_15S);
                                measureValues.Add(DataRow.W_15T);
                                measureValues.Add(DataRow.W_16S);
                                measureValues.Add(DataRow.W_16T);
                                measureValues.Add(DataRow.W_17S);
                                measureValues.Add(DataRow.W_17T);
                                measureValues.Add(DataRow.W1S_Peak);
                                measureValues.Add(DataRow.W1T_Peak);
                                measureValues.Add(DataRow.W2S_Peak);
                                measureValues.Add(DataRow.W2T_Peak);
                                measureValues.Add(DataRow.W3S_Peak);
                                measureValues.Add(DataRow.W3T_Peak);
                                measureValues.Add(DataRow.W4S_Peak);
                                measureValues.Add(DataRow.W4T_Peak);
                                measureValues.Add(DataRow.W5S_Peak);
                                measureValues.Add(DataRow.W5T_Peak);
                                measureValues.Add(DataRow.W6S_Peak);
                                measureValues.Add(DataRow.W6T_Peak);
                                measureValues.Add(DataRow.W7S_Peak);
                                measureValues.Add(DataRow.W7T_Peak);
                                measureValues.Add(DataRow.W8S_Peak);
                                measureValues.Add(DataRow.W8T_Peak);
                                measureValues.Add(DataRow.W9S_Peak);
                                measureValues.Add(DataRow.W9T_Peak);
                                measureValues.Add(DataRow.W10S_Peak);
                                measureValues.Add(DataRow.W10T_Peak);
                                measureValues.Add(DataRow.W11S_Peak);
                                measureValues.Add(DataRow.W11T_Peak);
                                measureValues.Add(DataRow.W12S_Peak);
                                measureValues.Add(DataRow.W12T_Peak);
                                measureValues.Add(DataRow.W13S_Peak);
                                measureValues.Add(DataRow.W13T_Peak);
                                measureValues.Add(DataRow.W14S_Peak);
                                measureValues.Add(DataRow.W14T_Peak);
                                measureValues.Add(DataRow.W15S_Peak);
                                measureValues.Add(DataRow.W15T_Peak);
                                measureValues.Add(DataRow.W16S_Peak);
                                measureValues.Add(DataRow.W16T_Peak);
                                measureValues.Add(DataRow.W17S_Peak);
                                measureValues.Add(DataRow.W17T_Peak);
                                #endregion

                                #region //columnName.add
                                columnName.Add("CheckFreq");
                                columnName.Add("FBL");
                                columnName.Add("DOF_Index");
                                columnName.Add("DOF_minus");
                                columnName.Add("DOF_plus");
                                columnName.Add("DOF_T");
                                columnName.Add("GDOF_Index");
                                columnName.Add("GDOF_minus");
                                columnName.Add("GDOF_plus");
                                columnName.Add("GDOF_T");
                                columnName.Add("EFL");
                                columnName.Add("W_1S");
                                columnName.Add("W_1T");
                                columnName.Add("W_2S");
                                columnName.Add("W_2T");
                                columnName.Add("W_3S");
                                columnName.Add("W_3T");
                                columnName.Add("W_4S");
                                columnName.Add("W_4T");
                                columnName.Add("W_5S");
                                columnName.Add("W_5T");
                                columnName.Add("W_6S");
                                columnName.Add("W_6T");
                                columnName.Add("W_7S");
                                columnName.Add("W_7T");
                                columnName.Add("W_8S");
                                columnName.Add("W_8T");
                                columnName.Add("W_9S");
                                columnName.Add("W_9T");
                                columnName.Add("W_10S");
                                columnName.Add("W_10T");
                                columnName.Add("W_11S");
                                columnName.Add("W_11T");
                                columnName.Add("W_12S");
                                columnName.Add("W_12T");
                                columnName.Add("W_13S");
                                columnName.Add("W_13T");
                                columnName.Add("W_14S");
                                columnName.Add("W_14T");
                                columnName.Add("W_15S");
                                columnName.Add("W_15T");
                                columnName.Add("W_16S");
                                columnName.Add("W_16T");
                                columnName.Add("W_17S");
                                columnName.Add("W_17T");
                                columnName.Add("W1S_Peak");
                                columnName.Add("W1T_Peak");
                                columnName.Add("W2S_Peak");
                                columnName.Add("W2T_Peak");
                                columnName.Add("W3S_Peak");
                                columnName.Add("W3T_Peak");
                                columnName.Add("W4S_Peak");
                                columnName.Add("W4T_Peak");
                                columnName.Add("W5S_Peak");
                                columnName.Add("W5T_Peak");
                                columnName.Add("W6S_Peak");
                                columnName.Add("W6T_Peak");
                                columnName.Add("W7S_Peak");
                                columnName.Add("W7T_Peak");
                                columnName.Add("W8S_Peak");
                                columnName.Add("W8T_Peak");
                                columnName.Add("W9S_Peak");
                                columnName.Add("W9T_Peak");
                                columnName.Add("W10S_Peak");
                                columnName.Add("W10T_Peak");
                                columnName.Add("W11S_Peak");
                                columnName.Add("W11T_Peak");
                                columnName.Add("W12S_Peak");
                                columnName.Add("W12T_Peak");
                                columnName.Add("W13S_Peak");
                                columnName.Add("W13T_Peak");
                                columnName.Add("W14S_Peak");
                                columnName.Add("W14T_Peak");
                                columnName.Add("W15S_Peak");
                                columnName.Add("W15T_Peak");
                                columnName.Add("W16S_Peak");
                                columnName.Add("W16T_Peak");
                                columnName.Add("W17S_Peak");
                                columnName.Add("W17T_Peak");
                                #endregion

                                Date_Time = DataRow.Date_Time;
                                string TrayLocation = LocationX + '-' + LocationY;
                                string QcResult = Class == "NG" ? "F" : "P";
                                string[] sArr = MfgOrder.Split(new char[3] { '-', '(', ')' });
                                string WoErpPrefix = sArr[0];
                                string WoErpNo = sArr[1];
                                int WoSeq = Convert.ToInt16(sArr[2]);
                                #endregion

                                #region //取得MoId
                                int MoId = -1;
                                int ModeId = -1;
                                int MoProcessId = -1;
                                string ProcessAlias = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId, a.ProcessAlias, a.MoProcessId
                                    , b.WoSeq, b.ModeId
                                    , c.WoErpPrefix, c.WoErpNo
                                    FROM MES.MoProcess a
                                    INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                    INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                    WHERE c.WoErpPrefix = @WoErpPrefix AND c.WoErpNo = @WoErpNo AND b.WoSeq = @WoSeq";
                                dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                                dynamicParameters.Add("WoErpNo", WoErpNo);
                                dynamicParameters.Add("WoSeq", WoSeq);
                                var MoIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (MoIdResult.Count() == 0)
                                {
                                    throw new SystemException("【MES】查無製令資訊!");
                                }
                                else
                                {
                                    foreach (var item2 in MoIdResult)
                                    {
                                        MoId = item2.MoId;
                                        ModeId = item2.ModeId;
                                        ProcessAlias = item2.ProcessAlias;
                                        MoProcessId = item2.MoProcessId;
                                    }
                                }
                                #endregion

                                #region //確認機台資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM MES.Machine a 
                                    WHERE a.MachineId = @MachineId";
                                dynamicParameters.Add("MachineId", MachineId);
                                var MachineResult = sqlConnection.Query(sql, dynamicParameters);
                                if (MachineResult.Count() <= 0) throw new SystemException("【MES】機台資料錯誤!!");
                                #endregion

                                #region //取得User資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT UserId
                                    FROM BAS.[User]
                                    WHERE UserNo = @UserNo";
                                dynamicParameters.Add("UserNo", UserNo);

                                var UserResult = sqlConnection.Query(sql, dynamicParameters);
                                if (UserResult.Count() <= 0) throw new SystemException("【MES】使用者資料錯誤!!");
                                int QcUserId = -1;
                                foreach (var item in UserResult)
                                {
                                    QcUserId = item.UserId;
                                }
                                #endregion

                                #region //確認條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.BarcodeId, b.BarcodeNo
	                                        FROM MES.Tray a
	                                        INNER JOIN MES.Barcode b ON a.TrayId = b.ParentBarcode
	                                        WHERE a.TrayNo = @TrayNo AND b.TrayLocation = @TrayLocation";
                                dynamicParameters.Add("TrayNo", TrayNo);
                                dynamicParameters.Add("TrayLocation", TrayLocation);

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (BarcodeResult.Count() <= 0) throw new Exception("【MES】TRAY條碼【" + TrayNo + "-" + TrayLocation + "】資料錯誤!!");
                                int BarcodeId = -1;
                                string BarcodeNo = "";
                                foreach (var item in BarcodeResult)
                                {
                                    BarcodeId = item.BarcodeId;
                                    BarcodeNo = item.BarcodeNo;

                                    if (item.BarcodeProcess == null)
                                    {
                                        bInput = true;
                                    }
                                    else
                                    {
                                        bInput = false;
                                    }
                                }
                                #endregion

                                #region //取得QcType
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT QcTypeId 
	                                        FROM QMS.QcType
                                            WHERE ModeId = @ModeId";
                                dynamicParameters.Add("ModeId", ModeId);
                                var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);
                                int QcTypeId = -1;
                                foreach (var item in QcTypeResult)
                                {
                                    QcTypeId = item.QcTypeId;
                                }
                                #endregion

                                QcTypeId = 20; //---------------------測試用

                                #region //取得預設SpreadsheetData
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";
                                var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);
                                string DefaultSpreadsheetData = "";
                                foreach (var item in DefaultFileIdResult)
                                {
                                    DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                                }
                                #endregion

                                #region //取得量測機台id、QmmDetailId
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QmmDetailId, a.MachineNumber, b.MachineName, b.MachineNo
                                        FROM  QMS.QmmDetail a
                                        INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                                        WHERE b.MachineId = @MachineId";
                                dynamicParameters.Add("MachineId", MachineId);
                                var QmmResult = sqlConnection.Query(sql, dynamicParameters);
                                if (QmmResult.Count() <= 0) throw new Exception("【MES】查無量測機台資訊");
                                int QmmDetailId = -1;
                                string MachineNumber = "";
                                foreach (var item in QmmResult)
                                {
                                    QmmDetailId = item.QmmDetailId;
                                    MachineNumber = item.MachineNumber;
                                }
                                #endregion

                                #region //INSERT MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcRecord (QcNoticeId, QcTypeId, InputType, MoId, MoProcessId, Remark, DefaultFileId, CurrentFileId, DefaultSpreadsheetData, CheckQcMeasureData, SupportAqFlag, SupportProcessFlag, ResolveFileJson
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordId
                                            VALUES (null, @QcTypeId, 'LotNo', @MoId, @MoProcessId, null, null, null, @DefaultSpreadsheetData, 'P', 'Y', 'N', '{""resolveFileInfo"":[]}'
                                            , @CreateDate, @LastModifiedDate, @UserId, @LastModifiedBy)";
                                dynamicParameters.Add("QcTypeId", QcTypeId);
                                dynamicParameters.Add("MoId", MoId);
                                dynamicParameters.Add("MoProcessId", MoProcessId);
                                dynamicParameters.Add("DefaultSpreadsheetData", DefaultSpreadsheetData);
                                dynamicParameters.Add("CreateDate", CreateDate);
                                dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                                dynamicParameters.Add("UserId", QcUserId);
                                dynamicParameters.Add("LastModifiedBy", QcUserId);

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                int rowsAffected = insertResult.Count();

                                int QcRecordId = -1;
                                foreach (var item in insertResult)
                                {
                                    QcRecordId = item.QcRecordId;
                                }
                                #endregion

                                #region //INSERT QMS.QcMeasureData，每一個數據都要存一筆 measureData
                                foreach (var tuple in measureValues.Zip(columnName, Tuple.Create))
                                {
                                    string measureValue = tuple.Item1;
                                    string QcItemName = tuple.Item2;

                                    #region //取得量測項目編號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT QcItemId, QcClassId, QcItemNo, QcItemName, QcItemDesc, QcItemType, QcType
		                                    FROM QMS.QcItem
		                                    WHERE QcItemName = @QcItemName";
                                    dynamicParameters.Add("QcItemName", QcItemName);

                                    var QcitemResult = sqlConnection.Query(sql, dynamicParameters);

                                    int QcItemId = -1;
                                    string QcItemNo = "";
                                    string QcItemDesc = "";
                                    foreach (var item in QcitemResult)
                                    {
                                        QcItemId = item.QcItemId;
                                        QcItemNo = item.QcItemNo;
                                        QcItemDesc = item.QcItemDesc;
                                    }
                                    #endregion

                                    string DesignValue = null;
                                    string UpperTolerance = null;
                                    string LowerTolerance = null;
                                    string ZAxis = null;
                                    string QcCycleTime = null;
                                    string CellHeader = null;
                                    string Row = null;
                                    string Remark = null;
                                    #region //INSERT QMS.QcMeasureData
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO QMS.QcMeasureData (QcRecordId, QcItemId, QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, ZAxis, BarcodeId, QmmDetailId, MeasureValue, QcResult, CauseId, CauseDesc, Cavity, LotNumber, MakeCount, CellHeader, Row, QcUserId, QcCycleTime, Remark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@QcRecordId, @QcItemId, @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance, @ZAxis, @BarcodeId, @QmmDetailId, @MeasureValue, @QcResult, @CauseId, @CauseDesc, @Cavity, @LotNumber, @MakeCount, @CellHeader, @Row, @QcUserId, @QcCycleTime, @Remark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            QcItemId,
                                            QcItemDesc,
                                            DesignValue,
                                            UpperTolerance,
                                            LowerTolerance,
                                            ZAxis,
                                            BarcodeId,
                                            QmmDetailId,
                                            MeasureValue = measureValue,
                                            QcResult,
                                            CauseId = QcResult == "P" ? (int?)null : 1311,
                                            CauseDesc = "",
                                            Cavity = "",
                                            LotNumber = Sup_LotCode ?? null,
                                            MakeCount = 1,
                                            CellHeader,
                                            Row,
                                            Remark,
                                            QcUserId,
                                            QcCycleTime,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var QcMeasureDataResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (QcMeasureDataResult.Count() == 0) throw new Exception("【MES】MTF數據新增失敗");
                                    rowsAffected += insertResult.Count();
                                    #endregion
                                }
                                #endregion

                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    woErpFullNo = WoErpPrefix + "-" + WoErpNo + "(" + WoSeq.ToString() + ")",
                                    barcodeNo = BarcodeNo,
                                    processAlias = ProcessAlias,
                                    bInput
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
        #endregion]
        #endregion

        #region //For SRM API
        #region //GetQcDeliveryInspectionForApi -- 取得供應商出貨檢驗單據資料(Api) -- Ann 2024-08-15
        public string GetQcDeliveryInspectionForApi(int QcDeliveryInspectionId, string PoErpFullNo, int SupplierId, int PoUserId, string SearchKey
            , string StartDate, string EndDate, string CompanyNo, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcDeliveryInspectionId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.QcRecordId, a.QcDiNo, a.PoDetailId, a.QcStatus, a.CreateUserName, a.UpdateUserName
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                        , c.PoSeq, c.PoMtlItemName, c.PoMtlItemSpec
                        , c.PoErpPrefix + '-' + c.PoErpNo + '(' + c.PoSeq + ')' PoErpFullNo
                        , d.PoErpPrefix, d.PoErpNo, d.Quantity
                        , e.CompanyNo
                        , f.StatusName
                        , g.MtlItemNo
                        , h.UserNo, h.UserName";
                    sqlQuery.mainTables =
                        @"FROM MES.QcDeliveryInspection a 
                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                        INNER JOIN SCM.PoDetail c ON a.PoDetailId = c.PoDetailId
                        INNER JOIN SCM.PurchaseOrder d ON c.PoId = d.PoId
                        INNER JOIN BAS.Company e ON d.CompanyId = e.CompanyId
                        INNER JOIN BAS.[Status] f ON f.StatusSchema = 'QcDeliveryInspection.QcStatus' AND f.StatusNo = a.QcStatus
                        INNER JOIN PDM.MtlItem g ON c.MtlItemId = g.MtlItemId
                        INNER JOIN BAS.[User] h ON d.PoUserId = h.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyNo", @" AND e.CompanyNo = @CompanyNo", CompanyNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcDeliveryInspectionId", @" AND a.QcDeliveryInspectionId = @QcDeliveryInspectionId", QcDeliveryInspectionId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpFullNo", @" AND (d.PoErpPrefix + '-' + d.PoErpNo) LIKE '%' + @PoErpFullNo + '%'", PoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND d.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoUserId", @" AND d.PoUserId = @PoUserId", PoUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND g.MtlItemNo LIKE '%' + @SearchKey + '%' OR c.PoMtlItemName LIKE '%' + @SearchKey + '%'", SearchKey);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcDeliveryInspectionId DESC";
                    sqlQuery.pageIndex = PageIndex <= 0 ? 1 : PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count() > 0 ? result.Select(x => x.TotalCount).FirstOrDefault() : 1,
                        recordsFiltered = result.Count() > 0 ? result.Select(x => x.TotalCount).FirstOrDefault() : 1,
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
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        #endregion

        #region //Model
        public class AqBarcode
        {
            public int? QcRecordId { get; set; }
            public int? BarcodeId { get; set; }
            public int? ConformUserId { get; set; }
            public int? ResponsibleDeptId { get; set; }
            public int? ResponsibleUserId { get; set; }
            public int? SubResponsibleDeptId { get; set; }
            public int? SubResponsibleUserId { get; set; }
            public int? ProgrammerUserId { get; set; }
            public int? MoProcessId { get; set; }
            public DateTime DocDate { get; set; }
            public string QcType { get; set; }
            public int? MoId { get; set; }
            public int? QcBarcodeId { get; set; }
            public int? DefectCauseId { get; set; }
            public int? GrDetailId { get; set; }
            public string DefectCauseDesc { get; set; }
            public int? SupplierId { get; set; }
        }

        public class QcNgCode
        {
            public int? BarcodeId { get; set; }
            public int? QcItemId { get; set; }
            public int? CauseId { get; set; }
            public string CauseDesc { get; set; }
        }
        #endregion

        #region 量測單寄送通知
        private void SendQcMail(SqlConnection sqlConnection, int QcRecordId, string WoErpFullNo, string MtlItemNo, string MtlItemName, string QcType, string QcTypeName
            , string Remark, string FileName, string ServerPath2, int MoId)
        {
            #region //量測完成自動計送信件
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
            dynamicParameters.Add("SettingSchema", "QcDataFinish");
            dynamicParameters.Add("SettingNo", "Y");

            var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
            if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
            #endregion

            foreach (var item in resultMailTemplate)
            {
                string mailSubject = item.MailSubject,
                    mailContent = HttpUtility.UrlDecode(item.MailContent);

                #region //Mail內容
                mailContent = mailContent.Replace("[QcRecordId]", QcRecordId.ToString());
                mailContent = mailContent.Replace("[FinishDate]", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                mailContent = mailContent.Replace("[WoErpFullNo]", WoErpFullNo);
                mailContent = mailContent.Replace("[MtlItemNo]", MtlItemNo);
                mailContent = mailContent.Replace("[MtlItemName]", MtlItemName);
                mailContent = mailContent.Replace("[QcType]", QcType + "(" + QcTypeName + ")");
                mailContent = mailContent.Replace("[Remark]", Remark);
                mailContent = mailContent.Replace("[QcDataHref]", "https://bm.zy-tech.com.tw/MesReport/IframeMeasurementRecord?QcRecordId=" + QcRecordId.ToString() + "");
                #endregion

                #region //設定附檔
                string QcFileNameCondition = "";
                List<string> FilePath = new List<string>();
                List<string> NewFilePath = new List<string>();
                List<MailFile> mailFiles = new List<MailFile>();
                if (FileName != "")
                {
                    #region //確認路徑格式
                    string date = FileName.Split('_')[2];
                    string year = "20" + date.Substring(0, 2);
                    string month = year + "-" + date.Substring(2, 2);
                    string folderPath = Path.Combine(ServerPath2, year, month, date, MoId.ToString());

                    if (FileName.Split('_').Length > 3)
                    {
                        if (FileName.Split('_')[3] == "BEFORE" || FileName.Split('_')[3] == "AFTER")
                        {
                            QcFileNameCondition = FileName.Split('_')[3];
                        }
                    }
                    #endregion

                    #region //抓取資料夾內所有檔案
                    if (Directory.Exists(folderPath))
                    {
                        FilePath = Directory.GetFiles(folderPath).ToList();
                    }
                    #endregion

                    #region //若為回火前、回火後的檔案，進行過濾
                    if (QcFileNameCondition.Length > 0)
                    {
                        if (QcFileNameCondition == "BEFORE") QcFileNameCondition = "回火前";
                        else if (QcFileNameCondition == "AFTER") QcFileNameCondition = "回火後";

                        foreach (string filePath in FilePath)
                        {
                            if (filePath.IndexOf(QcFileNameCondition) != -1)
                            {
                                NewFilePath.Add(filePath);
                            }
                        }
                    }

                    if (NewFilePath.Count() > 0) FilePath = NewFilePath;
                    #endregion

                    #region //組MailFile Modal
                    foreach (var filePath in FilePath)
                    {
                        MailFile mailFile = new MailFile()
                        {
                            FileName = Path.GetFileNameWithoutExtension(filePath),
                            FileExtension = Path.GetExtension(filePath),
                            FileContent = File.ReadAllBytes(filePath)
                        };

                        mailFiles.Add(mailFile);
                    }
                    #endregion
                }
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
                    TextBody = "-",
                    FileInfo = mailFiles,
                    QcFileFlag = "Y"
                };
                BaseHelper.MailSend(mailConfig);
                #endregion
            }
            #endregion
        }
        #endregion

        #region 品異單寄送
        private void SendIQcAqMail(SqlConnection sqlConnection, int QcRecordId, int AbnormalqualityId
            , string Remark, string FileName, string ServerPath2)
        {
            #region //自動送信件

            #region //取得MailServer
            dynamicParameters = new DynamicParameters();
            sql = @"select *
                    from BAS.MailServer
                    where CompanyId = @CompanyId";
            dynamicParameters.Add("CompanyId", CurrentCompany);

            var resultMailServer = sqlConnection.Query(sql, dynamicParameters);
            if (resultMailServer.Count() <= 0) throw new SystemException("MailServer設定錯誤!");
            var host = "";
            var account = "";
            var password = "";
            var port = 0;
            var sendMode = 0;

            foreach (var item in resultMailServer) {
                host = item.Host;
                account = item.Account;
                password = item.Password;
                port = Convert.ToInt32(item.Port);
                sendMode = Convert.ToInt32(item.SendMode);

            }
            #endregion

            #region //取得所有採購部門人員名單與信箱
            dynamicParameters = new DynamicParameters();
            sql = @"select b.DepartmentName, a.Job, a.UserName, a.Email 
                    from BAS.[User] a
                    inner join BAS.Department b on a.DepartmentId = b.DepartmentId
                    where a.UserStatus = 'F' and a.SystemStatus = 'A' and b.DepartmentId = @DepartmentId";

            dynamicParameters.Add("DepartmentId", 72);

            var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
            if (resultDepartment.Count() <= 0) throw new SystemException("部門設定錯誤!");
            string departmentMailString = string.Join(";", resultDepartment.Select(user =>
                                    $"{user.DepartmentName}{(string.IsNullOrEmpty(user.Job) ? "" : "-" + user.Job)}-{user.UserName}:{user.Email}"));
            #endregion

            #region //取得所有品保部門主管人員名單與信箱
            dynamicParameters = new DynamicParameters();
            sql = @"select b.DepartmentName, a.Job, a.UserName, a.Email , a.JobType
                    from BAS.[User] a
                    inner join BAS.Department b on a.DepartmentId = b.DepartmentId
                    where a.JobType = '管理制' and (a.DepartmentId = 1519 OR a.DepartmentId = 1520) and a.UserStatus = 'F'";

            var resultQcDepartment = sqlConnection.Query(sql, dynamicParameters);
            if (resultDepartment.Count() <= 0) throw new SystemException("部門設定錯誤!");
            string resultQcDepartmentstr = string.Join(";", resultQcDepartment.Select(user =>
                                    $"{user.DepartmentName}{(string.IsNullOrEmpty(user.Job) ? "" : "-" + user.Job)}-{user.UserName}:{user.Email}"));
            #endregion


            #region //取得品異單
            dynamicParameters = new DynamicParameters();
            
            sql = @"select * 
                    from QMS.Abnormalquality
                    where AbnormalqualityId = @AbnormalqualityId";
            dynamicParameters.Add("AbnormalqualityId", AbnormalqualityId);

            var resultAq = sqlConnection.Query(sql, dynamicParameters);
            if (resultAq.Count() <= 0) throw new SystemException("品異單錯誤!");
            var AbnormalqualityNo = "";
            foreach (var item in resultAq)
            {
                AbnormalqualityNo = item.AbnormalqualityNo;
                
            }
            #endregion


            #region //取得信件內容相關參數
            dynamicParameters = new DynamicParameters();
            sql = @"select a.QcGoodsReceiptId
                        , a.QcRecordId, a.GrDetailId, d.GrId, d.CompanyId
                        , e.MtlItemNo, e.MtlItemName, e.MtlItemSpec
						, (d.GrErpPrefix + '-' + d.GrErpNo + '-' + c.GrSeq) GrErpFullNo
                        , (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq) PoErpFullNo, (xd.UserNo + '-' + xd.UserName) PoConfidmUser
						, (xf.PrErpPrefix + '-' + xf.PrErpNo + '-' + xe.PrSequence) PrErpFullNo, (xg.UserNo + '-' + xg.UserName) PrConfidmUser, xg.UserId PrConfidmUserId
						FROM MES.QcGoodsReceipt a 
                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                        INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                        INNER JOIN SCM.GoodsReceipt d ON c.GrId = d.GrId
                        INNER JOIN PDM.MtlItem e ON c.MtlItemId = e.MtlItemId
                        INNER JOIN BAS.[User] f ON a.CreateBy = f.UserId
                        INNER JOIN BAS.[Status] g ON a.QcStatus = g.StatusNo AND g.StatusSchema = 'QcGoodsReceipt.QcStatus'
                        INNER JOIN SCM.Supplier h ON d.SupplierId = h.SupplierId
                        INNER JOIN BAS.Status i ON b.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'
                        INNER JOIN SCM.PoDetail xb ON (c.PoErpPrefix + '-' + c.PoErpNo + '-' + c.PoSeq) =  (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
						INNER JOIN SCM.PurchaseOrder xc ON xc.PoId = xb.PoId
						left JOIN BAS.[User] xd ON xc.PoUserId = xd.UserId
						left JOIN SCM.PurchaseRequisition xf ON (xb.PrErpPrefix + '-' + xb.PrErpNo) =  (xf.PrErpPrefix + '-' + xf.PrErpNo)
						left JOIN SCM.PrDetail xe ON xe.PrId = xf.PrId AND xb.PrSequence = xe.PrSequence
                        left JOIN BAS.[User] xg ON xf.UserId = xg.UserId
						where a.QcRecordId = @QcRecordId";
            dynamicParameters.Add("QcRecordId", QcRecordId);

            var resultQcRecord = sqlConnection.Query(sql, dynamicParameters);
            if (resultQcRecord.Count() <= 0) throw new SystemException("量測單據錯誤!");
            var GrErpFullNo = "";
            var MtlItemNo = "";
            var MtlItemName = "";
            var PrConfidmUserId = 0;


            foreach (var item in resultQcRecord)
            {
                GrErpFullNo = item.GrErpFullNo;
                MtlItemNo = item.MtlItemNo;
                MtlItemName = item.MtlItemName;
                PrConfidmUserId = item.PrConfidmUserId ?? -1;
            }
            #endregion

            #region //取得請購人信箱
            sql = @"select b.DepartmentName, a.Job, a.UserName, a.Email 
                    from BAS.[User] a
                    inner join BAS.Department b on a.DepartmentId = b.DepartmentId
                    where a.UserId = @UserId";

            dynamicParameters.Add("UserId", PrConfidmUserId);
            var resultPrUser = sqlConnection.Query(sql, dynamicParameters);
            string prUserMailString = "";
            if (resultPrUser.Count() > 0) {
                var prUser = resultPrUser.First();
                prUserMailString = $"{prUser.DepartmentName}{(string.IsNullOrEmpty(prUser.Job) ? "" : "-" + prUser.Job)}-{prUser.UserName}:{prUser.Email}";
            }
            
            #endregion

            string mailToString = departmentMailString + ";" + resultQcDepartmentstr + ";" + prUserMailString ;


            #region 內文
            string mailSubject = "進貨檢驗發生品異，請確認品異單內容";

            string mailContent = $@"
                                        各位採購部門、品保部門與請購同仁好：<br><br>
                                        
                                        本次進貨檢驗過程中，發現來料有品異情形，品異單已由系統自動開立，<br>
                                        敬請 請購人 依下列連結進行確認與處理，<br>
                                        如需協助，請與採購部門及品保部門聯繫。<br><br>
                                        
                                        🔗 品異單連結：<br>
                                        👉 <a href='https://bm.zy-tech.com.tw//Abnormalquality/AqJudgmentManagment?AbnormalqualityId={AbnormalqualityId}'>點我查看品異單</a><br><br>
                                        
                                        品異資訊如下：<br>
                                        ‧ 品異單號：{AbnormalqualityNo}<br>
                                        ‧ 異常量測單號：{QcRecordId}<br>
                                        ‧ 進貨單號：{GrErpFullNo}<br>
                                        ‧ 品號：{MtlItemNo}<br>
                                        ‧ 品名：{MtlItemName}<br>
                                        
                                        敬祝<br>
                                        工作順利<br>
                                        "; 
            #endregion

            #endregion

            #region //設定附檔
            string QcFileNameCondition = "";
                List<string> FilePath = new List<string>();
                List<string> NewFilePath = new List<string>();
                List<MailFile> mailFiles = new List<MailFile>();
                if (FileName != "")
                {
                    #region //確認路徑格式
                    string date = FileName.Split('_')[2];
                    string year = "20" + date.Substring(0, 2);
                    string month = year + "-" + date.Substring(2, 2);
                    string folderPath = Path.Combine(ServerPath2, year, month, date);

                    if (FileName.Split('_').Length > 3)
                    {
                        if (FileName.Split('_')[3] == "BEFORE" || FileName.Split('_')[3] == "AFTER")
                        {
                            QcFileNameCondition = FileName.Split('_')[3];
                        }
                    }
                    #endregion

                    #region //抓取資料夾內所有檔案
                    if (Directory.Exists(folderPath))
                    {
                        FilePath = Directory.GetFiles(folderPath).ToList();
                    }
                    #endregion

                    #region //若為回火前、回火後的檔案，進行過濾
                    if (QcFileNameCondition.Length > 0)
                    {
                        if (QcFileNameCondition == "BEFORE") QcFileNameCondition = "回火前";
                        else if (QcFileNameCondition == "AFTER") QcFileNameCondition = "回火後";

                        foreach (string filePath in FilePath)
                        {
                            if (filePath.IndexOf(QcFileNameCondition) != -1)
                            {
                                NewFilePath.Add(filePath);
                            }
                        }
                    }

                    if (NewFilePath.Count() > 0) FilePath = NewFilePath;
                    #endregion

                    #region //組MailFile Modal
                    foreach (var filePath in FilePath)
                    {
                        MailFile mailFile = new MailFile()
                        {
                            FileName = Path.GetFileNameWithoutExtension(filePath),
                            FileExtension = Path.GetExtension(filePath),
                            FileContent = File.ReadAllBytes(filePath)
                        };

                        mailFiles.Add(mailFile);
                    }
                    #endregion
                }
                #endregion

                #region //寄送Mail
                MailConfig mailConfig = new MailConfig
                {
                    Host = host,
                    Port = port,
                    SendMode = sendMode,
                    From = "企業管理平台:jmo-service@zy-tech.com.tw",
                    Subject = mailSubject,
                    Account = account,
                    Password = password,
                    MailTo = mailToString,
                    MailCc = "總經理-許智程:tim@zy-tech.com.tw;系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw",
                    //MailTo = "系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw",
                    //MailCc = "系統開發室-李宗宸:joe_li@zy-tech.com.tw;系統開發室-丁培文:pawun_ding@zy-tech.com.tw",
                    MailBcc = "",
                    HtmlBody = mailContent,
                    TextBody = "-",
                    FileInfo = mailFiles,
                    QcFileFlag = "Y"
                };
                BaseHelper.MailSend(mailConfig);
                #endregion
        }
        #endregion

        #region //FOR EIP API

        #region //GetQmmDetailEIP -- 取得量測機台資料 -- Ann 2023-04-12
        public string GetQmmDetailEIP(int QmmDetailId, int ShopId, string MachineNumber, int[] CustomerIds)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QmmDetailId, a.QcMachineModeId, a.MachineNumber, a.MachineId
                            , b.MachineDesc MachineName
                            FROM QMS.QmmDetail a
                            INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                            INNER JOIN MES.WorkShop c ON b.ShopId = c.ShopId
                            INNER JOIN BAS.Company y on c.CompanyId = y.CompanyId
                            INNER JOIN SCM.Customer z on y.CompanyId = z.CompanyId 
                            WHERE z.CustomerId IN @CustomerIds";
                    dynamicParameters.Add("CustomerIds", CustomerIds);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineNumber", @" AND a.MachineNumber = @MachineNumber", MachineNumber);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ShopId", @" AND c.ShopId = @ShopId", ShopId);

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
    }
}
