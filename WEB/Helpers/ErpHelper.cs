using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using Newtonsoft.Json.Linq;

namespace Helpers
{
    public class ErpHelper
    {
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static string sql = "";
        public JObject jsonResponse = new JObject();
        public DateTime LastModifiedDate = DateTime.Now;

        #region //GetDetailDesc 計算並取得領料備註(製令帶出使用)
        public string GetDetailDesc(string MtlItemNo, string InventoryNo, string MrErpPrefix, string MrErpNo, int MrId, int MtlItemId, int InventoryId, double MrDetailQty
            , SqlConnection mesSqlConnection, SqlConnection erpSqlConnection)
        {
            string DetailDesc = "";

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT a.MC007
                    FROM INVMC a
                    WHERE a.MC001 = @MC001
                    AND a.MC002 = @MC002";
            dynamicParameters.Add("MC001", MtlItemNo);
            dynamicParameters.Add("MC002", InventoryNo);

            var InventoryResult = erpSqlConnection.Query(sql, dynamicParameters);

            if (InventoryResult.Count() > 0)
            {
                foreach (var item2 in InventoryResult)
                {
                    #region //計算同料號同庫別，但未確認之領料單領料數量
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT ISNULL(SUM(a.TE005), 0) TE005
                            FROM MOCTE a
                            WHERE a.TE004 = @TE004
                            AND a.TE008 = @TE008
                            AND a.TE019 = 'N'
                            AND LTRIM(RTRIM(a.TE001)) + '_' + LTRIM(RTRIM(a.TE002)) != @MrErpFullNo";
                    dynamicParameters.Add("TE004", MtlItemNo);
                    dynamicParameters.Add("TE008", InventoryNo);
                    dynamicParameters.Add("MrErpFullNo", MrErpPrefix + "_" + MrErpNo);

                    var MtlItemQtyResult = erpSqlConnection.Query(sql, dynamicParameters);

                    double? TotalMtlItemQty = 0;
                    foreach (var item3 in MtlItemQtyResult)
                    {
                        TotalMtlItemQty = Convert.ToDouble(item3.TE005);
                    }
                    #endregion

                    #region //計算同領料單內，是否有別筆同品號同庫別之數量
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.Quantity
                            FROM MES.MrDetail a 
                            WHERE a.MrId = @MrId
                            AND a.MtlItemId = @MtlItemId
                            AND a.InventoryId = @InventoryId";
                    dynamicParameters.Add("MrId", MrId);
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    dynamicParameters.Add("InventoryId", InventoryId);

                    var MrDetailQtyResult = mesSqlConnection.Query(sql, dynamicParameters);

                    foreach (var item3 in MrDetailQtyResult)
                    {
                        TotalMtlItemQty += item3.Quantity;
                    }
                    #endregion

                    double? CurrentMtlItemQty = Convert.ToDouble(item2.MC007);
                    double? AvailabilityQty = CurrentMtlItemQty - TotalMtlItemQty/* - Convert.ToDouble(item.Quantity)*/;
                    if (AvailabilityQty >= 0)
                    {
                        DetailDesc = "可用量: " + AvailabilityQty;
                    }
                    else
                    {
                        DetailDesc = "庫存不足，需領用量: " + MrDetailQty;
                    }
                }
            }
            else
            {
                DetailDesc = "庫存不足，需領用量: " + MrDetailQty;
            }

            return DetailDesc;
        }
        #endregion

        #region //GetDetailDesc 計算並取得領料備註(拋轉段使用)
        public string GetDetailDesc(string MtlItemNo, string InventoryNo, string MrErpPrefix, string MrErpNo, int MrId, int MtlItemId, int InventoryId, double MrDetailQty, int MrDetailId
            , SqlConnection mesSqlConnection, SqlConnection erpSqlConnection)
        {
            string DetailDesc = "";

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT a.MC007
                    FROM INVMC a
                    WHERE a.MC001 = @MC001
                    AND a.MC002 = @MC002";
            dynamicParameters.Add("MC001", MtlItemNo);
            dynamicParameters.Add("MC002", InventoryNo);

            var InventoryResult = erpSqlConnection.Query(sql, dynamicParameters);

            if (InventoryResult.Count() > 0)
            {
                foreach (var item2 in InventoryResult)
                {
                    #region //計算同料號同庫別，但未確認之領料單領料數量
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT ISNULL(SUM(a.TE005), 0) TE005
                            FROM MOCTE a
                            WHERE a.TE004 = @TE004
                            AND a.TE008 = @TE008
                            AND a.TE019 = 'N'
                            AND LTRIM(RTRIM(a.TE001)) + '_' + LTRIM(RTRIM(a.TE002)) != @MrErpFullNo";
                    dynamicParameters.Add("TE004", MtlItemNo);
                    dynamicParameters.Add("TE008", InventoryNo);
                    dynamicParameters.Add("MrErpFullNo", MrErpPrefix + "_" + MrErpNo);

                    var MtlItemQtyResult = erpSqlConnection.Query(sql, dynamicParameters);

                    double? TotalMtlItemQty = 0;
                    foreach (var item3 in MtlItemQtyResult)
                    {
                        TotalMtlItemQty = Convert.ToDouble(item3.TE005);
                    }
                    #endregion

                    #region //計算同領料單內，是否有別筆同品號同庫別之數量
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MrDetailId, a.Quantity
                            FROM MES.MrDetail a 
                            WHERE a.MrId = @MrId
                            AND a.MtlItemId = @MtlItemId
                            AND a.InventoryId = @InventoryId
                            ORDER BY a.MrSequence ASC";
                    dynamicParameters.Add("MrId", MrId);
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    dynamicParameters.Add("InventoryId", InventoryId);

                    var MrDetailQtyResult = mesSqlConnection.Query(sql, dynamicParameters);
                    #endregion

                    double? CurrentMtlItemQty = Convert.ToDouble(item2.MC007);
                    double? AvailabilityQty = CurrentMtlItemQty - TotalMtlItemQty/* - Convert.ToDouble(item.Quantity)*/;

                    double? initialAvailability = AvailabilityQty;

                    foreach (var detail in MrDetailQtyResult)
                    {
                        double detailQty = Convert.ToDouble(detail.Quantity);

                        // 如果是當前處理的明細行，則計算並返回可用量
                        if (detail.MrDetailId == MrDetailId)
                        {
                            if (AvailabilityQty >= 0)
                            {
                                DetailDesc = "可用量: " + AvailabilityQty;
                            }
                            else
                            {
                                DetailDesc = "庫存不足，需領用量: " + MrDetailQty;
                            }
                            break; // 找到目標明細行後結束循環
                        }

                        // 不是當前處理的明細行，則減去其需求量，用於下一明細行的計算
                        AvailabilityQty -= detailQty;
                    }

                    if (string.IsNullOrEmpty(DetailDesc))
                    {
                        if (AvailabilityQty >= 0)
                        {
                            DetailDesc = "可用量: " + AvailabilityQty;
                        }
                        else
                        {
                            DetailDesc = "庫存不足，需領用量: " + MrDetailQty;
                        }
                    }
                }
            }
            else
            {
                DetailDesc = "庫存不足，需領用量: " + MrDetailQty;
            }

            return DetailDesc;
        }
        #endregion

        #region //單據確認檢核段
        #region //CheckMocYn 核對製令資料
        public string CheckMocYn(string FunctionName, string T1, string WoErpPrefix, string WoErpNo, SqlConnection sqlConnection)
        {
            string CheckMocFlag = "N";


            if (T1.Equals("Y"))
            {
                switch (FunctionName)
                {
                    case "MOCTI":
                    case "MOCTH":
                        #region //判斷庫別是否為存貨倉
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM MOCTA
                                WHERE TA001 = @WoErpPrefix
                                AND TA002 = @WoErpNo";
                        dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                        dynamicParameters.Add("WoErpNo", WoErpNo);
                        var MOCTAResult = sqlConnection.Query(sql, dynamicParameters);
                        if (MOCTAResult.Count() <= 0)
                        {
                            CheckMocFlag = "N";
                        }
                        else
                        {
                            CheckMocFlag = "Y";
                        }
                        #endregion
                        break;
                }

            }
            else
            {
                CheckMocFlag = "Y";
            }

            return CheckMocFlag;
        }
        #endregion

        #region //CheckInvmcYn 檢核品號是否有庫別庫存
        public int CheckInvmcYn(string p1, string p2,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            #region //搜尋該品號是否有庫別庫存
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT TOP 1 1
                      FROM INVMC
                      WHERE MC001 = @p1
                        AND MC002 = @p2";
            dynamicParameters.Add("p1", p1);
            dynamicParameters.Add("p2", p2);
            var Result = sqlConnection.Query(sql, dynamicParameters);
            if (Result.Count() <= 0)
            {
                #region //新增INVMC 
                dynamicParameters = new DynamicParameters();
                sql = @"INSERT INTO INVMC (COMPANY, CREATOR, USR_GROUP, 
                                           CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                           UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                           MC001, MC002)
                                   VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                           @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                           '','','','','',0,0,0,0,0,
                                           @p1 , @p2)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        p1,
                        p2
                    });
                rows += sqlConnection.Execute(sql, dynamicParameters);
                #endregion
            }
            #endregion

            return rows;
        }
        #endregion

        #region //CheckLotYn 檢核批號管理
        public string CheckLotYn(string t1, string t2)
        {
            string CheckLotFlag = "N";
            if (!t1.Equals("N") && t2.Equals("******************"))
            {
                CheckLotFlag = "N";
            }
            else
            {

                CheckLotFlag = "Y";
            }
            return CheckLotFlag;
        }
        #endregion

        #region //CheckLotQty 檢核批號庫存
        public string CheckLotQty(string T1, string T2, string T3, double T4, double T5, SqlConnection sqlConnection)
        {
            string CheckLotQtyFlag = "N";

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT SUM(MF008*MF010)  AS P_QTY, SUM(MF008*MF014) AS K_QTY
                      FROM INVMF
                     WHERE MF001 = @T1 
                       AND MF002 = @T2 
                       AND MF007 = @T3 ";
            dynamicParameters.Add("T1", T1);
            dynamicParameters.Add("T2", T2);
            dynamicParameters.Add("T3", T3);

            var CheckLotResult = sqlConnection.Query(sql, dynamicParameters);
            if (CheckLotResult.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
            double P_QTY = Convert.ToDouble(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).P_QTY);
            double K_QTY = Convert.ToDouble(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).K_QTY);
            if (P_QTY + T4 >= 0 && K_QTY + T5 >= 0)
            {
                CheckLotQtyFlag = "Y";
            }
            else
            {
                CheckLotQtyFlag = "N";
            }

            return CheckLotQtyFlag;
        }
        #endregion

        #region //CheckStoQty 檢查庫別庫存
        public string CheckStoQty(string T1, string T2, double T3, double T4, SqlConnection sqlConnection, int D_O5)
        {
            string CheckStoFlag = "N";

            if (D_O5 > 0) return "Y";

            //CMSMA.MA024 1代表沒啟用包裝單位
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MA024
                      FROM CMSMA";
            var CMSResult = sqlConnection.Query(sql, dynamicParameters);
            if (CMSResult.Count() <= 0) throw new SystemException("找不到共用參數檔!");
            string MA024 = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MA024;

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MC007  AS S_QTY, MC014  AS K_QTY 
                      FROM INVMC
                     WHERE MC001 = @T1 
                       AND MC002 = @T2 ";
            dynamicParameters.Add("T1", T1);
            dynamicParameters.Add("T2", T2);

            var CheckStoResult = sqlConnection.Query(sql, dynamicParameters);
            if (CheckStoResult.Count() <= 0) throw new SystemException("找不到品號【" + T1 + "】庫別【" + T2 + "】庫存資訊!!!");
            double S_QTY = Convert.ToDouble(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).S_QTY);
            double K_QTY = Convert.ToDouble(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).K_QTY);
            if (!MA024.Equals("1"))
            {
                if (S_QTY + T3 >= 0 && K_QTY + T4 >= 0)
                {
                    CheckStoFlag = "Y";
                }
                else
                {
                    CheckStoFlag = "N";
                }
            }
            else
            {
                if (S_QTY + T3 >= 0)
                {
                    CheckStoFlag = "Y";
                }
                else
                {
                    CheckStoFlag = "N";
                }
            }


            return CheckStoFlag;
        }
        #endregion

        #region //CheckMocQty 核對製令超領
        public string CheckMocQty(string T1, string T2, string T3, double T4, string T5, SqlConnection sqlConnection, string BomMtlItemNo, decimal BomMtlItemQty, string D_A1, string D_A2, int D_O5, double TD006)
        {
            //T5 = 控制超領
            //D_O5 = 庫存影響 1:增 -1:減
            //T4 = 此筆單身領料數量
            //TD006 = 領退料單套數
            string CheckMocFlag = "N";

            if (T5.Equals("Y"))
            {
                if (D_O5 < 0)
                {
                    #region //檢核領料單是否超領
                    #region //檢查此料號是否為替代料，若是則改以原元件檢查是否超領
                    if (BomMtlItemNo != T1 && BomMtlItemQty != 0)
                    {
                        #region //確認此製令單身是否有其他此原元件替代料，並加總計算已領用量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TE030
                                FROM MOCTE 
                                WHERE TE011 = @T2
                                AND TE012 = @T3
                                AND TE027 = @BomMtlItemNo
                                AND TE004 != @BomMtlItemNo
                                AND TE019 = 'Y'";
                        dynamicParameters.Add("T2", T2);
                        dynamicParameters.Add("T3", T3);
                        dynamicParameters.Add("BomMtlItemNo", BomMtlItemNo);

                        var MOCTEResult = sqlConnection.Query(sql, dynamicParameters);

                        decimal sumBomQty = 0;
                        foreach (var item in MOCTEResult)
                        {
                            sumBomQty += item.TE030;
                        }
                        #endregion

                        #region //檢核原元件是否超領
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TB004, TB005
                                FROM MOCTB 
                                WHERE TB003 = @BomMtlItemNo
                                AND TB001 = @T2
                                AND TB002 = @T3";
                        dynamicParameters.Add("BomMtlItemNo", BomMtlItemNo);
                        dynamicParameters.Add("T2", T2);
                        dynamicParameters.Add("T3", T3);

                        var BomMtlItemMOCTBResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in BomMtlItemMOCTBResult)
                        {
                            if (BomMtlItemQty > item.TB004)
                            {
                                throw new SystemException("料號【" + T1 + "】超過原元件需領用量，請確認!!");
                            }
                            else
                            {
                                CheckMocFlag = "Y";
                                return CheckMocFlag;
                            }
                        }
                        #endregion
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TB004-TB005 AS MOCQTY 
                        FROM MOCTB
                        WHERE TB001 = @T2 
                        AND TB002 = @T3 
                        AND TB003 = @T1 ";
                    dynamicParameters.Add("T1", T1);
                    dynamicParameters.Add("T2", T2);
                    dynamicParameters.Add("T3", T3);

                    var CheckStoResult = sqlConnection.Query(sql, dynamicParameters);
                    if (CheckStoResult.Count() <= 0) throw new SystemException("找不到製令領料資訊!!!");
                    decimal MOCQTY = Convert.ToDecimal(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MOCQTY);
                    if (MOCQTY >= Convert.ToDecimal(T4))
                    {
                        CheckMocFlag = "Y";
                    }
                    else
                    {
                        CheckMocFlag = "N";
                    }
                    #endregion

                }
                else
                {
                    double allowQty = 0;

                    #region //取得製令相關資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TB004, a.TB005
                            , b.TA015, b.TA016, b.TA017, b.TA018
                            FROM MOCTB a 
                            INNER JOIN MOCTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                            WHERE a.TB001 = @T2 
                            AND a.TB002 = @T3 
                            AND a.TB003 = @T1";
                    dynamicParameters.Add("T1", T1);
                    dynamicParameters.Add("T2", T2);
                    dynamicParameters.Add("T3", T3);

                    var MOCTBResult = sqlConnection.Query(sql, dynamicParameters);

                    if (MOCTBResult.Count() <= 0) throw new SystemException("查詢製令相關資料時錯誤!!");

                    double TB004 = 0; //需領用量
                    double TB005 = 0; //已領用量
                    double TA015 = 0; //預計生產量
                    double TA016 = 0; //已領套數
                    double TA017 = 0; //已生產量
                    double TA018 = 0; //報廢數量
                    foreach (var item in MOCTBResult)
                    {
                        TB004 = Convert.ToDouble(item.TB004);
                        TB005 = Convert.ToDouble(item.TB005);
                        TA015 = Convert.ToDouble(item.TA015);
                        TA016 = Convert.ToDouble(item.TA016);
                        TA017 = Convert.ToDouble(item.TA017);
                        TA018 = Convert.ToDouble(item.TA018);
                    }
                    #endregion
                    #region //檢核退料單套數是否超退
                    //超退套數檢核公式: 本次退套數 MOCTD.TD006 > 已領套數 MOCTA.TA016 - ( 已生產量MOCTA.TA017 + 報廢數量MOCTA.TA018)

                    //計算可退數量(套數)
                    allowQty = TA016 - (TA017 + TA018);
                    if (TD006 > 0 && TD006 > allowQty)
                    {
                        return "N";
                    }

                    #region //檢核退料單單身數量是否超退
                    //超退檢核公式: 本次退料量 MOCTE.TE005 > 已領用量 MOCTB.TB005 - (需領用量MOCTB.TB004 * (已生產量MOCTA.TA017 + 報廢數量MOCTA.TA018) / 預計產量MOCTA.TA015))

                    //計算可超退數量
                    allowQty = TB005 - (TB004 * (TA017 + TA018) / TA015);

                    if (T4 > allowQty)
                    {
                        CheckMocFlag = "N";
                    }
                    else
                    {
                        CheckMocFlag = "Y";
                    }
                    #endregion
                    #endregion
                }
            }
            else
            {
                CheckMocFlag = "Y";
            }


            return CheckMocFlag;
        }
        #endregion

        #region //CheckLocationQty 檢核儲位庫存
        public string CheckLocationQty(string MtlItemNo, string InventoryNo, string StorageLocation, double Quantity, SqlConnection sqlConnection)
        {
            string CheckLotQtyFlag = "N";

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MM005
                    FROM INVMM
                    WHERE MM001 = @MtlItemNo
                    AND MM002 = @InventoryNo
                    AND MM003 = @StorageLocation";
            dynamicParameters.Add("MtlItemNo", MtlItemNo);
            dynamicParameters.Add("InventoryNo", InventoryNo);
            dynamicParameters.Add("StorageLocation", StorageLocation);

            var CheckLocationResult = sqlConnection.Query(sql, dynamicParameters);
            if (CheckLocationResult.Count() <= 0) throw new SystemException("找不到儲位庫存紀錄資訊!!!");

            foreach (var item in CheckLocationResult)
            {
                double inventoryQty = Convert.ToDouble(item.MM005);
                if (Quantity > inventoryQty)
                {
                    CheckLotQtyFlag = "N";
                }
                else
                {
                    CheckLotQtyFlag = "Y";
                }
            }

            return CheckLotQtyFlag;
        }
        #endregion
        #endregion

        #region //單據確認異動段
        #region //InsertInvmm 更新品號庫別儲位批號檔(INVMM)
        public int InsertInvmm(string p1, string p2, string p3, string p4, double p5,
                               double p6, string p7,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string LastTradeDate = "";
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");


            if (FunctionName == "MOCTH")
            {
                LastTradeDate = "MM008";
            }
            else
            {
                LastTradeDate = "MM009";
            }

            //CMSMA.MA024 1代表沒啟用包裝單位
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MA024
                      FROM CMSMA";
            var CMSResult = sqlConnection.Query(sql, dynamicParameters);
            if (CMSResult.Count() <= 0) throw new SystemException("找不到共用參數檔!");
            string MA024 = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MA024;

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MM001 
                      FROM INVMM 
                     WHERE MM001 = @p1
                       AND MM002 = @p2
                       AND MM003 = @p3 
                       AND MM004 = @p4";
            dynamicParameters.Add("p1", p1);
            dynamicParameters.Add("p2", p2);
            dynamicParameters.Add("p3", p3.Length > 0 ? p3 : "##########");
            dynamicParameters.Add("p4", p4.Length > 0 ? p4 : "####################");
            var chkInvmmResult = sqlConnection.Query(sql, dynamicParameters);

            if (chkInvmmResult.Count() > 0)
            {
                dynamicParameters = new DynamicParameters();
                sql = @"UPDATE INVMM
                           SET MODIFIER =@MODIFIER,
                               MODI_DATE= @MODI_DATE,
                               MODI_TIME= @MODI_TIME,
                               MODI_AP= @MODI_AP,
                               MODI_PRID= @MODI_PRID,
                               FLAG = CASE 
                                  WHEN FLAG >= 999 THEN 1
                                  ELSE FLAG + 1
                               END,
                               MM005 = MM005 + @p5,
                               MM006 = MM006 + @p6
                         WHERE MM001 = @p1
                           AND MM002 = @p2
                           AND MM003 = @p3
                           AND MM004 = @p4";
                dynamicParameters.AddDynamicParams(
                new
                {
                    MODIFIER = UserNo,
                    MODI_DATE = dateNow,
                    MODI_TIME = timeNow,
                    MODI_AP = BaseHelper.ClientComputer(),
                    MODI_PRID = "BM",
                    p1,
                    p2,
                    p3 = p3.Length > 0 ? p3 : "##########",
                    p4 = p4.Length > 0 ? p4 : "####################",
                    p5,
                    p6
                });

                rows = sqlConnection.Execute(sql, dynamicParameters);
            }
            else
            {
                dynamicParameters = new DynamicParameters();
                sql = @"INSERT INTO INVMM (COMPANY, CREATOR, USR_GROUP, 
                                           CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                           UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                           MM001, MM002, MM003, MM004, MM005,
                                           MM006, " + LastTradeDate + @")
                                   VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                           @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                           '','','','','',0,0,0,0,0,
                                           @p1 , @p2 , @p3 , @p4 , @p5 , 
                                           @p6 , @p7)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        p1,
                        p2,
                        p3 = p3.Length > 0 ? p3 : "##########",
                        p4 = p4.Length > 0 ? p4 : "####################",
                        p5,
                        p6,
                        p7
                    });

                rows = sqlConnection.Execute(sql, dynamicParameters);
            }

            if (rows <= 0) throw new SystemException("執行確認新增【品號庫別儲位批號檔】時發生錯誤!! 方法名稱:【InsertInvmm】");

            return rows;
        }
        #endregion

        #region //InsertInvmef 新增庫存批號檔(INVME、INVMF)
        public int InsertInvmef(string p1, string p2, string p3, string p4, string p5, string p6,
                                string p7, string p8, string p9, int p10, string p11, double p12, string WoErpPrefix, string WoErpNo,
                                string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");


            dynamicParameters = new DynamicParameters();
            sql = @"SELECT ME001 
                      FROM INVME 
                     WHERE ME001 = @p1 
                       AND ME002 = @p2";
            dynamicParameters.Add("p1", p1);
            dynamicParameters.Add("p2", p2);
            var InvmeResult = sqlConnection.Query(sql, dynamicParameters);

            dynamicParameters = new DynamicParameters();
            if (InvmeResult.Count() > 0)
            {
                #region //新增INVMF
                sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, 
                                           CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                           UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                           MF001, MF002, MF003, MF004, MF005,
                                           MF006, MF007, MF008, MF009, MF010, MF013)
                                   VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                           @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                           '','','','','',0,0,0,0,0,
                                           @p1, @p2, @p3, @p4, @p5, 
                                           @p8, @p9, @p10, @p11, @p12, @MF013)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        p1,
                        p2,
                        p3,
                        p4,
                        p5,
                        p8,
                        p9,
                        p10,
                        p11,
                        p12,
                        MF013 = WoErpPrefix + WoErpNo
                    });
                var insertResult1 = sqlConnection.Execute(sql, dynamicParameters);

                if (insertResult1 <= 0) throw new SystemException("執行確認新增【新增庫存批號檔】時發生錯誤!! 方法名稱:【InsertInvmef】");
                #endregion
            }
            else
            {
                #region //新增INVME
                sql = @"INSERT INTO INVME (COMPANY, CREATOR, USR_GROUP, 
                                           CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                           UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                           ME001, ME002, ME003, ME005, ME006, 
                                           ME007, ME008, ME009, ME010)
                                   VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                           @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                           '','','','','',0,0,0,0,0,
                                           @p1, @p2, @p3, @p4, @p5,
                                           'N', @ME008, @p6, @p7)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        p1,
                        p2,
                        p3,
                        p4,
                        p5,
                        ME008 = WoErpPrefix + WoErpNo,
                        p6,
                        p7
                    });
                var insertResult1 = sqlConnection.Execute(sql, dynamicParameters);

                if (insertResult1 <= 0) throw new SystemException("執行確認新增【新增庫存批號檔】時發生錯誤!! 方法名稱:【InsertInvmef】");
                #endregion

                #region //新增INVMF
                sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, 
                                           CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                           UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                           MF001, MF002, MF003, MF004, MF005,
                                           MF006, MF007, MF008, MF009, MF010, MF013)
                                   VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                           @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                           '','','','','',0,0,0,0,0,
                                           @p1, @p2, @p3, @p4, @p5, 
                                           @p8, @p9, @p10, @p11, @p12, @MF013)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        p1,
                        p2,
                        p3,
                        p4,
                        p5,
                        p8,
                        p9,
                        p10,
                        p11,
                        p12,
                        MF013 = WoErpPrefix + WoErpNo
                    });
                var insertResult2 = sqlConnection.Execute(sql, dynamicParameters);

                if (insertResult2 <= 0) throw new SystemException("執行確認新增【新增庫存批號檔】時發生錯誤!! 方法名稱:【InsertInvmef】");
                #endregion
            }
            return rows;
        }
        #endregion

        #region //InsertInvla 新增庫存庫存異動檔(INVLA)
        public int InsertInvla(string p1, string p2, int p3, string p4, string p5,
                               string p6, string p7, string p8, double p9, double p10,
                               double p11, string p12, string p13, string p14, double p15,
                               string p16, string WoErpPrefix, string WoErpNo, string Supplier,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string p8Date = "";

            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");


            #region //判斷成本碼
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT b1.TA006,b1.MQ003 WMQ003,a.MQ010,a.MQ003,a.MQ011
                    FROM CMSMQ a
                    OUTER APPLY(
                        SELECT x.MQ003,x1.TA006,x1.TA001 FROM CMSMQ x
                        INNER JOIN MOCTA x1 on x.MQ001 = x1.TA001
                        WHERE x1.TA001 = @WoErpPrefix AND x1.TA002 = @WoErpNo
                    ) b1
                    WHERE a.MQ001 = @p4";
            //AND a.TE004 = b.TA006
            //AND c.MQ003 = '52'";
            dynamicParameters.Add("p4", p4);
            dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
            dynamicParameters.Add("WoErpNo", WoErpNo);
            var LA015Result = sqlConnection.Query(sql, dynamicParameters);
            if (LA015Result.Count() > 0)
            {
                string TA006 = "";
                string WMQ003 = "";
                string MQ003 = "";
                string MQ011 = "";
                foreach (var la in LA015Result)
                {
                    TA006 = la.TA006;
                    WMQ003 = la.WMQ003;
                    MQ003 = la.MQ003;
                    MQ011 = la.MQ011;
                }
                if (MQ003.Equals("54") || MQ003.Equals("55") || MQ003.Equals("56") || MQ003.Equals("57"))
                {
                    if (p1.Equals(TA006))
                    {
                        p13 = "r";
                    }
                    else
                    {
                        p13 = MQ011;
                    }
                }
                else if (MQ003.Equals("59") || MQ003.Equals("5A"))
                {
                    if (WMQ003.Equals("52"))
                    {
                        p13 = "R";
                    }
                    else
                    {
                        p13 = MQ011;
                    }
                }
                else
                {
                    //之後挪料單為5C 待確認邏輯
                    p13 = MQ011;
                }
            }
            else
            {
                throw new SystemException("核單邏輯異常, 請通知系統開發室!");
            }
            #endregion

            #region //判斷庫別是否為存貨倉
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MC004 
                      FROM CMSMC
                     WHERE MC001 = @p7
                       AND MC004 != '1'";
            dynamicParameters.Add("p7", p7);
            var CMSMCResult = sqlConnection.Query(sql, dynamicParameters);
            if (CMSMCResult.Count() > 0)
            {
                p10 = 0;
                p11 = 0;
            }
            #endregion

            if (FunctionName == "MOCTH")
            {
                p8Date = WoErpPrefix + "-" + WoErpNo + " " + Supplier;
            }
            else if (FunctionName.Equals("INVTF"))
            {
                p8Date = Supplier;
            }
            else
            {
                p8Date = WoErpPrefix + "-" + WoErpNo;
            }

            #region //取得本國幣別
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

            var CMSMAResult = sqlConnection.Query(sql, dynamicParameters);

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

            var CMSMFResult = sqlConnection.Query(sql, dynamicParameters);

            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

            int unitDecimal = 0;
            int amountDecimal = 0;
            foreach (var item in CMSMFResult)
            {
                unitDecimal = Convert.ToInt32(item.MF005);
                amountDecimal = Convert.ToInt32(item.MF006);
            }
            #endregion

            #region //新增INVLA
            dynamicParameters = new DynamicParameters();
            sql = @"INSERT INTO INVLA (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       LA001, LA004, LA005, LA006, LA007,
                                       LA008, LA009, LA010, LA011, LA012,
                                       LA013, LA014, LA015, LA016, LA021,
                                       LA022)
                                OUTPUT INSERTED.LA001
                               VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @p1, @p2, @p3, @p4, @p5,
                                       @p6, @p7, @p8, @p9, @p10,
                                       @p11, @p12, @p13, @p14, @p15,
                                       @p16)";
            dynamicParameters.AddDynamicParams(
                new
                {
                    COMPANY = ErpNo,
                    CREATOR = UserNo,
                    USR_GROUP,
                    CREATE_DATE = dateNow,
                    MODIFIER = "",
                    MODI_DATE = "",
                    FLAG = 1,
                    CREATE_TIME = timeNow,
                    CREATE_AP = BaseHelper.ClientComputer(),
                    CREATE_PRID = "BM",
                    p1,
                    p2,
                    p3,
                    p4,
                    p5,
                    p6,
                    p7,
                    p8 = p8Date,
                    p9,
                    p10 = Math.Round(p10, unitDecimal),
                    p11 = Math.Round(p11, amountDecimal),
                    p12,
                    p13,
                    p14,
                    p15,
                    p16
                });

            rows = sqlConnection.Execute(sql, dynamicParameters);
            #endregion

            if (rows <= 0) throw new SystemException("執行確認新增【庫存庫存異動檔】時發生錯誤!! 方法名稱:【InsertInvla】");

            return rows;
        }
        #endregion

        #region //InsertInvld 品號專案管理檔(INVLD)
        public int InsertInvld(string p1, string p2, string p3, int p4, string p5,
                               string p6, string p7, string p8, string p9, double p10,
                               double p11, double p12, string p13, double p14,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            #region //新增INVLD
            dynamicParameters = new DynamicParameters();
            sql = @"INSERT INTO INVLD (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       LD001, LD002, LD003, LD004, LD005,
                                       LD006, LD007, LD008, LD009, LD010,
                                       LD011, LD012, LD013, LD014)
                               VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @p1, @p2, @p3, @p4, @p5,
                                       @p6, @p7, @p8, @p9, @p10,
                                       @p11, @p12, @p13, @p14)";
            dynamicParameters.AddDynamicParams(
                new
                {
                    COMPANY = ErpNo,
                    CREATOR = UserNo,
                    USR_GROUP,
                    CREATE_DATE = dateNow,
                    MODIFIER = "",
                    MODI_DATE = "",
                    FLAG = 1,
                    CREATE_TIME = timeNow,
                    CREATE_AP = BaseHelper.ClientComputer(),
                    CREATE_PRID = "BM",
                    p1,
                    p2,
                    p3,
                    p4,
                    p5,
                    p6,
                    p7,
                    p8,
                    p9,
                    p10,
                    p11,
                    p12,
                    p13,
                    p14
                });

            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行確認新增【品號專案管理檔】時發生錯誤!! 方法名稱:【InsertInvld】");
            #endregion
            return rows;
        }
        #endregion

        #region //UpdateInvmb 更新品號基本檔(INVMB)
        public int UpdateInvmb(string p1, double p2, double p3,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            #region //取得本國幣別
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

            var CMSMAResult = sqlConnection.Query(sql, dynamicParameters);

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

            var CMSMFResult = sqlConnection.Query(sql, dynamicParameters);

            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

            int unitDecimal = 0;
            int amountDecimal = 0;
            foreach (var item in CMSMFResult)
            {
                unitDecimal = Convert.ToInt32(item.MF005);
                amountDecimal = Convert.ToInt32(item.MF006);
            }
            #endregion

            #region //更新INVMB
            dynamicParameters = new DynamicParameters();
            sql = @"UPDATE INVMB
                       SET MODIFIER = @MODIFIER,
                           MODI_DATE= @MODI_DATE,
                           MODI_TIME= @MODI_TIME,
                           MODI_AP= @MODI_AP,
                           MODI_PRID= @MODI_PRID,
                           FLAG = CASE 
                                WHEN FLAG >= 999 THEN 1
                                ELSE FLAG + 1
                           END,
                           MB064 = MB064 + @p2,
                           MB065 = MB065 + @p3
                     WHERE MB001 = @p1";
            dynamicParameters.AddDynamicParams(
            new
            {
                MODIFIER = UserNo,
                MODI_DATE = dateNow,
                MODI_TIME = timeNow,
                MODI_AP = BaseHelper.ClientComputer(),
                MODI_PRID = "BM",
                p1,
                p2,
                p3 = Math.Round(p3, amountDecimal)
            });

            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行確認更新【品號基本檔】時發生錯誤!! 方法名稱:【UpdateInvmb】");
            #endregion

            return rows;
        }
        #endregion

        #region //UpdateInvmc 更新庫別檔(INVMC)
        public int UpdateInvmc(string p1, double p2, double p3, string p4, string ReceiptDate,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP, string ErpPrefix)
        {
            double a = Math.Round(p3, 6);
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");
            string MC012 = "";
            string MC013 = "";

            #region //確認單據 出庫日/入庫日 是否需要更新
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT LTRIM(RTRIM(MQ012)) MQ012, LTRIM(RTRIM(MQ013)) MQ013
                    FROM CMSMQ
                    WHERE MQ001 = @ErpPrefix";
            dynamicParameters.Add("ErpPrefix", ErpPrefix);

            var CMSMQResult = sqlConnection.Query(sql, dynamicParameters);

            if (CMSMQResult.Count() <= 0) throw new SystemException("確認單據設定參數錯誤，單據【" + ErpPrefix + "】");

            string inputFlag = "";
            string outputFlag = "";
            foreach (var item in CMSMQResult)
            {
                inputFlag = item.MQ012;
                outputFlag = item.MQ013;
            }
            #endregion

            #region //取得原本出/入庫日
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT MC012, MC013
                    FROM INVMC
                    WHERE MC001 = @p1
                    AND MC002 = @p4";
            dynamicParameters.Add("p1", p1);
            dynamicParameters.Add("p4", p4);
            var result = sqlConnection.Query(sql, dynamicParameters);
            foreach (var item in result)
            {
                MC012 = item.MC012;
                MC013 = item.MC013;
            }
            #endregion

            #region //處理出入庫日格式
            if (inputFlag == "Y" && ReceiptDate != "")
            {
                if (MC012 == "")
                {
                    MC012 = ReceiptDate;
                }

                if (Convert.ToInt32(ReceiptDate) > Convert.ToInt32(MC012))
                {
                    MC012 = ReceiptDate;
                }
            }

            if (outputFlag == "Y" && ReceiptDate != "")
            {
                if (MC013 == "")
                {
                    MC013 = ReceiptDate;
                }

                if (Convert.ToInt32(ReceiptDate) > Convert.ToInt32(MC013))
                {
                    MC013 = ReceiptDate;
                }
            }
            #endregion

            #region //取得本國幣別
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

            var CMSMAResult = sqlConnection.Query(sql, dynamicParameters);

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

            var CMSMFResult = sqlConnection.Query(sql, dynamicParameters);

            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

            int unitDecimal = 0;
            int amountDecimal = 0;
            foreach (var item in CMSMFResult)
            {
                unitDecimal = Convert.ToInt32(item.MF005);
                amountDecimal = Convert.ToInt32(item.MF006);
            }
            #endregion

            if (result.Count() > 0)
            {
                #region //UPDATE INVMC 
                dynamicParameters = new DynamicParameters();
                sql = @"UPDATE INVMC
                           SET MODIFIER =@MODIFIER,
                               MODI_DATE= @MODI_DATE,
                               MODI_TIME= @MODI_TIME,
                               MODI_AP= @MODI_AP,
                               MODI_PRID= @MODI_PRID,
                               FLAG = CASE 
                                  WHEN FLAG >= 999 THEN 1
                                  ELSE FLAG + 1
                               END,
                               MC007 = MC007 + @p2,
                               MC008 = MC008 + @p3,
                               MC012 = @MC012,
                               MC013 = @MC013
                         WHERE MC001 = @p1
                           AND MC002 = @p4";
                dynamicParameters.AddDynamicParams(
                new
                {
                    MODIFIER = UserNo,
                    MODI_DATE = dateNow,
                    MODI_TIME = timeNow,
                    MODI_AP = BaseHelper.ClientComputer(),
                    MODI_PRID = "BM",
                    p1,
                    p2,
                    p3 = Math.Round(p3, amountDecimal),
                    p4,
                    MC012,
                    MC013
                });

                rows = sqlConnection.Execute(sql, dynamicParameters);
                #endregion
            }
            else
            {
                #region //INSERT INVMC
                dynamicParameters = new DynamicParameters();
                sql = @"INSERT INTO INVMC (COMPANY, CREATOR, USR_GROUP, 
                                CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                MC001, MC002, MC003, MC004, MC005, MC006, MC007, MC008, MC009, MC010, MC011,
                                MC012, MC013, MC014, MC015)
                        VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                @MC001, @MC002, @MC003, @MC004, @MC005, @MC006, @MC007, @MC008, @MC009, @MC010, @MC011,
                                @MC012, @MC013, @MC014, @MC015)";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        MC001 = p1,
                        MC002 = p4,
                        MC003 = "",
                        MC004 = 0.000,
                        MC005 = 0.000,
                        MC006 = 0.000,
                        MC007 = p2,
                        MC008 = Math.Round(p3, amountDecimal),
                        MC009 = 0.000,
                        MC010 = 0.00000,
                        MC011 = "",
                        MC012,
                        MC013,
                        MC014 = 0.000,
                        MC015 = ""
                    });
                rows = sqlConnection.Execute(sql, dynamicParameters);
                #endregion
            }

            if (rows <= 0) throw new SystemException("執行確認更新【品號庫別檔】時發生錯誤!! 方法名稱:【UpdateInvmc】");

            return rows;
        }
        #endregion

        #region //InsertInvlf 新增儲位批號異動明細資料(INVLF)
        public int InsertInvlf(string p1, string p2, string p3, string p4, string p5,
                               string p6, string p7, int p8, string p9, string p10,
                               double p11, double p12, string p13, string p14, string p15,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            #region //新增INVLF
            dynamicParameters = new DynamicParameters();
            sql = @"INSERT INTO INVLF (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       LF001, LF002, LF003, LF004, LF005,
                                       LF006, LF007, LF008, LF009, LF010,
                                       LF011, LF012, LF013)
                               VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @p1, @p2, @p3, @p4, @p5,
                                       @p6, @p7, @p8, @p9, @p10,
                                       @p11, @p12 , @p13)";
            dynamicParameters.AddDynamicParams(
                new
                {
                    COMPANY = ErpNo,
                    CREATOR = UserNo,
                    USR_GROUP,
                    CREATE_DATE = dateNow,
                    MODIFIER = "",
                    MODI_DATE = "",
                    FLAG = 1,
                    CREATE_TIME = timeNow,
                    CREATE_AP = BaseHelper.ClientComputer(),
                    CREATE_PRID = "BM",
                    p1,
                    p2,
                    p3,
                    p4,
                    p5,
                    p6 = p6.Length <= 0 ? "##########" : p6,
                    p7 = p7.Length <= 0 ? "####################" : p7,
                    p8,
                    p9,
                    p10,
                    p11,
                    p12,
                    p13 = p14 + p15
                });
            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行確認新增【儲位批號異動明細】時發生錯誤!! 方法名稱:【InsertInvlf】");
            #endregion

            return rows;
        }
        #endregion

        #region //InsertInvme 新增品號批號單頭(INVME)
        public int InsertInvme(string p1, string p2, string p3, string p4, string p5,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;
            string dateNow = DateTime.Now.ToString("yyyyMMdd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            dynamicParameters = new DynamicParameters();
            sql = @"SELECT ME001 
                      FROM INVME 
                     WHERE ME001 = @p1
                       AND ME002 = @p2";
            dynamicParameters.Add("p1", p1);
            dynamicParameters.Add("p2", p2);
            var chkInvmmResult = sqlConnection.Query(sql, dynamicParameters);

            if (chkInvmmResult.Count() <= 0)
            {
                dynamicParameters = new DynamicParameters();
                sql = @"INSERT INTO INVME (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       ME001, ME002, ME005, ME006, ME007, 
                                       MM009)
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @p1 , @p2 , @p3 , @p4 , 'N',
                                       @p5 )";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        COMPANY = ErpNo,
                        CREATOR = UserNo,
                        USR_GROUP,
                        CREATE_DATE = dateNow,
                        MODIFIER = "",
                        MODI_DATE = "",
                        FLAG = 1,
                        CREATE_TIME = timeNow,
                        CREATE_AP = BaseHelper.ClientComputer(),
                        CREATE_PRID = "BM",
                        p1,
                        p2,
                        p3,
                        p4,
                        p5
                    });

                rows = sqlConnection.Execute(sql, dynamicParameters);
            }

            if (rows <= 0) throw new SystemException("執行確認新增【品號批號單頭】時發生錯誤!! 方法名稱:【InsertInvme】");

            return rows;
        }
        #endregion

        #region //InsertInvmf 新增品號批號單身(INVMF)
        public int InsertInvmf(string p1, string p2, string p3, string p4, string p5,
                               string p6, string p7, string p8, string p9, double p10,
                               double p11,
                               string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;

            dynamicParameters = new DynamicParameters();
            sql = @"INSERT INTO INVLA (COMPANY, CREATOR, USR_GROUP, CREATE_DATE,
                                       MODIFIER, MODI_DATE, FLAG, UDF01, UDF02,
                                       UDF03, UDF04, UDF05, UDF06, UDF07,
                                       UDF08, UDF09, UDF10,
                                       MF001, MF002, MF003, MF004, MF005, 
                                       MF006, MF007, MF008, MF009, MF010, 
                                       MF014)
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       CONVERT(varchar(100),GETDATE(),112),@CREATOR,CONVERT(varchar(100),GETDATE(),112),0,
                                       '','','','','',0,0,0,0,0,
                                       @p1 , @p2 , @p3 , @p4 , @p5,
                                       @p6 , @p7 , @p8 , @p9 , @p10,
                                       @p11)";
            dynamicParameters.AddDynamicParams(
                new
                {
                    COMPANY = ErpNo,
                    CREATOR = UserNo,
                    USR_GROUP,
                    p1,
                    p2,
                    p3,
                    p4,
                    p5,
                    p6,
                    p7,
                    p8,
                    p9,
                    p10,
                    p11
                });

            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行確認新增【品號批號單身】時發生錯誤!! 方法名稱:【InsertInvmf】");

            return rows;
        }
        #endregion
        #endregion

        #region //單據反確認異動段
        #region //DelInvmef 刪除批號檔(INVME、INVMF)
        public int DelInvmef(string p1, string p2, string p3, string p4, string p5,
                             string p6,
                             string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;

            dynamicParameters = new DynamicParameters();
            sql = @"DELETE INVMF
                     WHERE MF001 = @p1
                       AND MF002 = @p2
                       AND MF003 = @p3
                       AND MF004 = @p4
                       AND MF005 = @p5
                       AND MF006 = @p6";
            dynamicParameters.AddDynamicParams(
            new
            {
                p1,
                p2,
                p3,
                p4,
                p5,
                p6
            });

            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行反確認刪除【批號檔】時發生錯誤!! 方法名稱:【DelInvmef】");

            return rows;
        }
        #endregion

        #region //DelInvla 刪除庫存異動檔(INVLA)
        public int DelInvla(string p1, string p2, int p3, string p4, string p5,
                            string p6,
                            string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;

            dynamicParameters = new DynamicParameters();
            sql = @"DELETE INVLA
                     WHERE LA001 = @p1
                       AND LA004 = @p2
                       AND LA005 = @p3
                       AND LA006 = @p4
                       AND LA007 = @p5
                       AND LA008 = @p6";
            dynamicParameters.AddDynamicParams(
            new
            {
                p1,
                p2,
                p3,
                p4,
                p5,
                p6
            });

            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行反確認刪除【庫存異動檔】時發生錯誤!! 方法名稱:【DelInvla】");

            return rows;
        }
        #endregion

        #region //DelInvlf 刪除批號異動明細資料(INVLF)
        public int DelInvlf(string p1, string p2, string p3, string p4, string p5,
                            string p6, string p7,
                            string UserNo, string FunctionName, string ErpNo, SqlConnection erpSqlConnection, string USR_GROUP)
        {
            int rows = 0;

            dynamicParameters = new DynamicParameters();
            sql = @"DELETE INVLF
                     WHERE LF001 = @p1
                       AND LF002 = @p2
                       AND LF003 = @p3
                       AND LF004 = @p4
                       AND LF005 = @p5
                       AND LF006 = @p6
                       AND LF007 = @p7";
            dynamicParameters.AddDynamicParams(
            new
            {
                p1,
                p2,
                p3,
                p4,
                p5,
                p6 = p6.Length <= 0 ? "##########" : p6,
                p7 = p7.Length <= 0 ? "####################" : p7,
            });

            rows = erpSqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行反確認刪除【批號異動明細資料】時發生錯誤!! 方法名稱:【DelInvlf】");

            return rows;
        }
        #endregion

        #region //DelInvld 刪除品號專案管理檔(INVLD)
        public int DelInvld(string ProjectCode, string MtlItemNo, string TransactionDate, int IncOrDec, string TransactionPrefix,
                            string TransactionNo, string TransactionSeq,
                            string UserNo, string FunctionName, string ErpNo, SqlConnection sqlConnection, string USR_GROUP)
        {
            int rows = 0;

            dynamicParameters = new DynamicParameters();
            sql = @"DELETE INVLD
                     WHERE LD001 = @ProjectCode
                       AND LD002 = @MtlITemNo
                       AND LD003 = @TransactionDate
                       AND LD004 = @IncOrDec
                       AND LD005 = @TransactionPrefix
                       AND LD006 = @TransactionNo
                       AND LD007 = @TransactionSeq";
            dynamicParameters.AddDynamicParams(
            new
            {
                ProjectCode,
                MtlItemNo,
                TransactionDate,
                IncOrDec,
                TransactionPrefix,
                TransactionNo,
                TransactionSeq
            });

            rows = sqlConnection.Execute(sql, dynamicParameters);

            if (rows <= 0) throw new SystemException("執行反確認刪除【品號專案管理檔】時發生錯誤!! 方法名稱:【DelInvld】");

            return rows;
        }
        #endregion
        #endregion

        #region //進貨單相關API
        #region //UpdateTransferGoodsReceipt -- 拋轉進貨單據 -- Ann 2024-03-11
        public string UpdateTransferGoodsReceipt(int GrId, SqlConnection sqlConnection, SqlConnection sqlConnection2, int UserId, int CompanyId)
        {
            try
            {
                string companyNo = "";
                int rowsAffected = 0;

                List<GoodsReceipt> goodsReceipts = new List<GoodsReceipt>();
                List<GrDetail> grDetails = new List<GrDetail>();

                #region //更新BM進貨單資料
                JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, -1, sqlConnection, sqlConnection2, UserId));
                if (updateGrDetailResult["status"].ToString() != "success")
                {
                    throw new SystemException(updateGrDetailResult["msg"].ToString());
                }
                #endregion

                #region //公司別資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.ErpNo, a.ErpDb
                        FROM BAS.Company a
                        WHERE CompanyId = @CompanyId";
                dynamicParameters.Add("CompanyId", CompanyId);

                var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                foreach (var item in resultCompany)
                {
                    companyNo = item.ErpNo;
                }
                #endregion

                #region //取得BM進貨單頭資料
                sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo
                        , CASE WHEN LEN(a.ReceiptDate) > 0 THEN FORMAT(CAST(a.ReceiptDate as date), 'yyyy-MM-dd') ELSE null END AS ReceiptDate
                        , CASE WHEN LEN(a.ReceiptDate) > 0 THEN FORMAT(CAST(a.ReceiptDate as date), 'yyyyMMdd') ELSE '' END AS ErpReceiptDate
                        , a.SupplierId, a.SupplierNo, a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus
                        , a.DocDate, CASE WHEN LEN(a.DocDate) > 0 THEN FORMAT(CAST(a.DocDate as date), 'yyyyMMdd') ELSE '' END AS ErpDocDate
                        , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                        , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity
                        , CASE WHEN FORMAT(a.InvoiceDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(a.InvoiceDate as date), 'yyyy-MM-dd') END AS InvoiceDate
                        , CASE WHEN LEN(a.InvoiceDate) > 0 THEN FORMAT(CAST(a.InvoiceDate as date), 'yyyyMMdd') ELSE '' END AS ErpInvoiceDate
                        , a.OrigPreTaxAmount, a.ApplyYYMM, a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount
                        , a.TaxOffset, a.PackageQuantity, a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo
                        , a.EbcEdition, a.EbcFlag, a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                        , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                        FROM SCM.GoodsReceipt a 
                        WHERE a.CompanyId = @CompanyId
                        AND a.GrId = @GrId";
                dynamicParameters.Add("CompanyId", CompanyId);
                dynamicParameters.Add("GrId", GrId);

                goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                if (goodsReceipts[0].ConfirmStatus != "N") throw new SystemException("進貨單目前狀態不可拋轉!!");

                #region //整理單頭日期格式
                string DocDate = ((DateTime)goodsReceipts.FirstOrDefault().DocDate).ToString("yyyy-MM-dd");
                string ReceiptDate = ((DateTime)goodsReceipts.FirstOrDefault().ReceiptDate).ToString("yyyy-MM-dd");
                string InvoiceDate = goodsReceipts.FirstOrDefault().InvoiceDate == null ? "" : ((DateTime)goodsReceipts.FirstOrDefault().InvoiceDate).ToString("yyyyMMdd");
                #endregion
                #endregion

                #region //取得BM進貨單身資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.GrDetailId, a.GrId, a.GrErpPrefix, a.GrErpNo, a.GrSeq, a.MtlItemId, a.MtlItemNo, a.GrMtlItemName, a.GrMtlItemSpec
                        , a.ReceiptQty, a.UomId, a.UomNo, a.InventoryId, a.InventoryNo, a.LotNumber, a.PoErpPrefix, a.PoErpNo, a.PoSeq
                        , a.AcceptanceDate, CASE WHEN LEN(a.AcceptanceDate) > 0 THEN FORMAT(CAST(a.AcceptanceDate as date), 'yyyyMMdd') ELSE '' END AS ErpAcceptanceDate
                        , a.AcceptQty, a.AvailableQty, a.ReturnQty, a.OrigUnitPrice, a.OrigAmount, a.OrigDiscountAmt, a.TrErpPrefix, a.TrErpNo, a.TrSeq
                        , a.ReceiptExpense, a.DiscountDescription, a.PaymentHold, a.Overdue, a.QcStatus, a.ReturnCode, a.ConfirmStatus, a.CloseStatus
                        , a.ReNewStatus, a.Remark, a.InventoryQty, a.SmallUnit
                        , CASE WHEN FORMAT(a.AvailableDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(a.AvailableDate as date), 'yyyy-MM-dd') END AS AvailableDate
                        , CASE WHEN LEN(a.AvailableDate) > 0 THEN FORMAT(CAST(a.AvailableDate as date), 'yyyyMMdd') ELSE '' END AS ErpAvailableDate
                        , CASE WHEN FORMAT(a.ReCheckDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(a.ReCheckDate as date), 'yyyy-MM-dd') END AS ReCheckDate
                        , CASE WHEN LEN(a.ReCheckDate) > 0 THEN FORMAT(CAST(a.ReCheckDate as date), 'yyyyMMdd') ELSE '' END AS ErpReCheckDate
                        , a.ConfirmUser, a.ApvErpPrefix, a.ApvErpNo
                        , a.ApvSeq, a.ProjectCode, a.ExpenseEntry, a.PremiumAmountFlag, a.OrigPreTaxAmt, a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.ReceiptPackageQty
                        , a.AcceptancePackageQty, a.ReturnPackageQty, a.PremiumAmount, a.PackageUnit, a.ReserveTaxCode, a.OrigPremiumAmount, a.AvailableUomId
                        , a.AvailableUomNo, a.OrigCustomer, a.ApproveStatus, a.EbcErpPreNo, a.EbcEdition, a.ProductSeqAmount, a.MtlItemType, a.Loaction
                        , a.TaxRate, a.MultipleFlag, a.GrErpStatus, a.TaxCode, a.DiscountRate, a.DiscountPrice, a.QcType, a.TransferStatus, a.TransferDate
                        , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                        FROM SCM.GrDetail a
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                grDetails = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();
                #endregion

                #region //確認ERP進貨單單頭資料是否正確
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT LTRIM(RTRIM(TG013)) TG013
                        FROM PURTG
                        WHERE TG001 = @GrErpPrefix
                        AND TG002 = @GrErpNo";
                dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);
                dynamicParameters.Add("GrErpNo", goodsReceipts[0].GrErpNo);

                var PURTGResult = sqlConnection2.Query(sql, dynamicParameters);

                foreach (var item in PURTGResult)
                {
                    if (item.TG013 != "N") throw new SystemException("進貨單目前狀態不可拋轉!!");
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
                    DateTime DocDateDateTime = Convert.ToDateTime(DocDate.ToString().Split('-')[0] + "-" + DocDate.ToString().Split('-')[1]);
                    int compare = DocDateDateTime.CompareTo(erpDateTime);
                    if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + goodsReceipts[0].DocDate.ToString() + ")之單據!!");
                }
                #endregion

                #region //取得目前使用者資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT UserNo
                        FROM BAS.[User]
                        WHERE UserId = @UserId";
                dynamicParameters.Add("UserId", UserId);

                var UserResult = sqlConnection.Query(sql, dynamicParameters);

                string UserNo = "";
                foreach (var item in UserResult)
                {
                    UserNo = item.UserNo;
                }
                #endregion

                #region //審核ERP權限
                string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "", "");
                #endregion

                #region //查詢廠別
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT LTRIM(RTRIM(MB001)) MB001 FROM CMSMB";
                var CMSMBResult = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMBResult.Count() <= 0) throw new SystemException("找不到廠別!!");
                if (CMSMBResult.Count() > 1) throw new SystemException("廠別數有多個，請與資訊人員確認!!");

                string factoryNo = "";
                foreach (var item in CMSMBResult)
                {
                    factoryNo = item.MB001;
                }
                #endregion

                #region //新增/修改單頭PURTG資料
                if (PURTGResult.Count() > 0)
                {
                    #region //UPDATE PURTG 進貨單單頭
                    foreach (var item in goodsReceipts)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PURTG SET
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                FLAG += 1,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                TG003 = @ReceiptDate,
                                TG005 = @SupplierNo,
                                TG007 = @CurrencyCode,
                                TG008 = @Exchange,
                                TG009 = @InvoiceType,
                                TG010 = @TaxType,
                                TG011 = @InvoiceNo,
                                TG014 = @DocDate,
                                TG016 = @Remark,
                                TG017 = @TotalAmount,
                                TG018 = @DeductAmount,
                                TG019 = @OrigTax,
                                TG020 = @ReceiptAmount,
                                TG021 = @SupplierName,
                                TG023 = @DeductType,
                                TG026 = @Quantity,
                                TG027 = @InvoiceDate,
                                TG028 = @OrigPreTaxAmount,
                                TG029 = @ApplyYYMM,
                                TG030 = @TaxRate,
                                TG031 = @PreTaxAmount,
                                TG032 = @TaxAmount,
                                TG033 = @PaymentTerm,
                                TG034 = @PoErpPrefix,
                                TG035 = @PoErpNo,
                                TG041 = @PremiumAmount,
                                TG058 = @TaxCode,
                                TG059 = @TradeTerm,
                                TG067 = @ContactUser
                                WHERE TG001 = @GrErpPrefix
                                AND TG002 = @GrErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = DateTime.Now.ToString("yyyyMMdd"),
                            MODI_TIME = DateTime.Now.ToString("HH:mm:ss"),
                            MODI_AP = UserNo + "PC",
                            MODI_PRID = "BM",
                            ReceiptDate = item.ErpReceiptDate,
                            item.SupplierNo,
                            item.CurrencyCode,
                            item.Exchange,
                            item.InvoiceType,
                            item.TaxType,
                            item.InvoiceNo,
                            DocDate = item.ErpDocDate,
                            item.Remark,
                            item.TotalAmount,
                            item.DeductAmount,
                            item.OrigTax,
                            item.ReceiptAmount,
                            item.SupplierName,
                            item.DeductType,
                            item.Quantity,
                            InvoiceDate = item.ErpInvoiceDate,
                            item.OrigPreTaxAmount,
                            item.ApplyYYMM,
                            item.TaxRate,
                            item.PreTaxAmount,
                            item.TaxAmount,
                            item.PaymentTerm,
                            item.PoErpPrefix,
                            item.PoErpNo,
                            item.PremiumAmount,
                            item.TaxCode,
                            item.TradeTerm,
                            item.ContactUser,
                            item.GrErpPrefix,
                            item.GrErpNo,
                        });

                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    }
                    #endregion

                    #region //刪除原本單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"DELETE PURTH
                            WHERE TH001 = @GrErpPrefix
                            AND TH002 = @GrErpNo";
                    dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);
                    dynamicParameters.Add("GrErpNo", goodsReceipts[0].GrErpNo);

                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion
                }
                else
                {
                    #region //取號流程
                    string GrErpNo = "";
                    DateTime referenceTime = DateTime.ParseExact(goodsReceipts[0].ErpDocDate, "yyyyMMdd", CultureInfo.InvariantCulture);

                    #region //單據設定
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                            FROM CMSMQ a
                            WHERE a.COMPANY = @CompanyNo
                            AND a.MQ001 = @ErpPrefix";
                    dynamicParameters.Add("CompanyNo", companyNo);
                    dynamicParameters.Add("ErpPrefix", goodsReceipts[0].GrErpPrefix);

                    var resultDocSetting = sqlConnection2.Query(sql, dynamicParameters);
                    if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                    string encode = "";
                    int yearLength = -1;
                    int lineLength = -1;
                    foreach (var item in resultDocSetting)
                    {
                        encode = item.MQ004; //編碼方式
                        yearLength = Convert.ToInt32(item.MQ005); //年碼數
                        lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                    }
                    #endregion

                    #region //單號取號
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TG002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                        FROM PURTG
                                        WHERE TG001 = @ErpPrefix";
                    dynamicParameters.Add("ErpPrefix", goodsReceipts[0].GrErpPrefix);

                    #region //編碼方式
                    string dateFormat = "";
                    switch (encode)
                    {
                        case "1": //日編
                            dateFormat = new string('y', yearLength) + "MMdd";
                            sql += @" AND RTRIM(LTRIM(TG002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                            GrErpNo = referenceTime.ToString(dateFormat);
                            break;
                        case "2": //月編
                            dateFormat = new string('y', yearLength) + "MM";
                            sql += @" AND RTRIM(LTRIM(TG002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                            GrErpNo = referenceTime.ToString(dateFormat);
                            break;
                        case "3": //流水號
                            break;
                        case "4": //手動編號
                            break;
                        default:
                            throw new SystemException("編碼方式錯誤!");
                    }
                    #endregion

                    var currentNum = sqlConnection2.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                    GrErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                    goodsReceipts[0].GrErpNo = GrErpNo;
                    #endregion
                    #endregion

                    #region //INSERT PURTG 進貨單單頭
                    foreach (var item in goodsReceipts)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PURTG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID
                                , TG001, TG002, TG003, TG004, TG005, TG006, TG007, TG008, TG009, TG010, TG011, TG012, TG013, TG014, TG015, TG016, TG017
                                , TG018, TG019, TG020, TG021, TG022, TG023, TG024, TG025, TG026, TG027, TG028, TG029, TG030, TG031, TG032, TG033, TG034
                                , TG035, TG036, TG037, TG038, TG039, TG040, TG041, TG042, TG043, TG044, TG045, TG046, TG047, TG048, TG049, TG050, TG051
                                , TG052, TG053, TG054, TG055, TG056, TG057, TG058, TG059, TG060, TG061, TG062, TG063, TG064, TG065, TG066, TG067, TG068
                                , TG500, TG550, TG069, TG070, TG071, TG072, TG073, TG074, TG075, TG076, TG077)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID
                                , @TG001, @TG002, @TG003, @TG004, @TG005, @TG006, @TG007, @TG008, @TG009, @TG010, @TG011, @TG012, @TG013, @TG014, @TG015, @TG016, @TG017
                                , @TG018, @TG019, @TG020, @TG021, @TG022, @TG023, @TG024, @TG025, @TG026, @TG027, @TG028, @TG029, @TG030, @TG031, @TG032, @TG033, @TG034
                                , @TG035, @TG036, @TG037, @TG038, @TG039, @TG040, @TG041, @TG042, @TG043, @TG044, @TG045, @TG046, @TG047, @TG048, @TG049, @TG050, @TG051
                                , @TG052, @TG053, @TG054, @TG055, @TG056, @TG057, @TG058, @TG059, @TG060, @TG061, @TG062, @TG063, @TG064, @TG065, @TG066, @TG067, @TG068
                                , @TG500, @TG550, @TG069, @TG070, @TG071, @TG072, @TG073, @TG074, @TG075, @TG076, @TG077)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                COMPANY = companyNo,
                                CREATOR = UserNo,
                                USR_GROUP,
                                CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                FLAG = 1,
                                CREATE_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                CREATE_AP = UserNo + "PC",
                                CREATE_PRID = "BM",
                                TG001 = item.GrErpPrefix,
                                TG002 = item.GrErpNo,
                                TG003 = item.ErpReceiptDate,
                                TG004 = factoryNo,
                                TG005 = item.SupplierNo,
                                TG006 = item.SupplierSo,
                                TG007 = item.CurrencyCode,
                                TG008 = item.Exchange,
                                TG009 = item.InvoiceType,
                                TG010 = item.TaxType,
                                TG011 = item.InvoiceNo,
                                TG012 = item.PrintCnt,
                                TG013 = item.ConfirmStatus,
                                TG014 = item.ErpDocDate,
                                TG015 = item.RenewFlag,
                                TG016 = item.Remark,
                                TG017 = item.TotalAmount,
                                TG018 = item.DeductAmount,
                                TG019 = item.OrigTax,
                                TG020 = item.ReceiptAmount,
                                TG021 = item.SupplierName,
                                TG022 = item.UiNo,
                                TG023 = item.DeductType,
                                TG024 = item.ObaccoAndLiquorFlag,
                                TG025 = item.RowCnt,
                                TG026 = item.Quantity,
                                TG027 = InvoiceDate,
                                TG028 = item.OrigPreTaxAmount,
                                TG029 = item.ApplyYYMM,
                                TG030 = item.TaxRate,
                                TG031 = item.PreTaxAmount,
                                TG032 = item.TaxAmount,
                                TG033 = item.PaymentTerm,
                                TG034 = item.PoErpPrefix,
                                TG035 = item.PoErpNo,
                                TG036 = item.PrepaidErpPrefix,
                                TG037 = item.PrepaidErpNo,
                                TG038 = item.OffsetAmount,
                                TG039 = item.TaxOffset,
                                TG040 = item.PackageQuantity,
                                TG041 = item.PremiumAmount,
                                TG042 = item.ApproveStatus,
                                TG043 = item.InvoiceFlag,
                                TG044 = item.ReserveTaxCode,
                                TG045 = item.SendCount,
                                TG046 = item.OrigPremiumAmount,
                                TG047 = item.EbcErpPreNo,
                                TG048 = item.EbcEdition,
                                TG049 = item.EbcFlag,
                                TG050 = item.FromDocType,
                                TG051 = "",
                                TG052 = "",
                                TG053 = 0.000000,
                                TG054 = 0.000000,
                                TG055 = item.NoticeFlag,
                                TG056 = "",
                                TG057 = "",
                                TG058 = item.TaxCode,
                                TG059 = item.TradeTerm,
                                TG060 = "",
                                TG061 = "",
                                TG062 = "",
                                TG063 = item.DetailMultiTax,
                                TG064 = item.TaxExchange,
                                TG065 = 0,
                                TG066 = 0,
                                TG067 = item.ContactUser,
                                TG068 = "",
                                TG500 = "",
                                TG550 = "",
                                TG069 = "",
                                TG070 = "",
                                TG071 = "",
                                TG072 = "",
                                TG073 = "",
                                TG074 = "",
                                TG075 = "",
                                TG076 = "",
                                TG077 = ""
                            });
                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    }
                    #endregion
                }
                #endregion

                #region //新增單身PURTH資料
                foreach (var item in grDetails)
                {
                    #region //確認ERP採購單身結案碼
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(TD016)) TD016
                            , LTRIM(RTRIM(TD018)) TD018
                            FROM PURTD
                            WHERE TD001 = @PoErpPrefix
                            AND TD002 = @PoErpNo
                            AND TD003 = @PoSeq";
                    dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                    dynamicParameters.Add("PoErpNo", item.PoErpNo);
                    dynamicParameters.Add("PoSeq", item.PoSeq);

                    var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                    if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】資料錯誤!!");

                    foreach (var item2 in PURTDResult)
                    {
                        if (item.ConfirmStatus != "Y")
                        {
                            if (item2.TD016 != "N") throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】已結案!!");
                        }
                        if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】狀態錯誤!!");
                    }
                    #endregion

                    #region //確認ERP品號相關資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                            , LTRIM(RTRIM(MB030)) MB030
                            , LTRIM(RTRIM(MB031)) MB031
                            , LTRIM(RTRIM(MB022)) MB022
                            FROM INVMB
                            WHERE MB001 = @MB001";
                    dynamicParameters.Add("MB001", item.MtlItemNo);

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
                            int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
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
                            int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                            if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                            #endregion
                        }
                        #endregion

                        #region //確認此品號是否需要批號管理
                        if (item2.MB022 == "Y" || item2.MB022 == "T")
                        {
                            if (item.LotNumber.Length <= 0) throw new SystemException("品號【" + item.MtlItemNo + "】需設定批號!!");
                        }
                        else
                        {
                            if (item.LotNumber.Length > 0) throw new SystemException("品號【" + item.MtlItemNo + "】不需設定批號!!");
                        }
                        #endregion
                    }
                    #endregion

                    #region //整理單身日期格式
                    string AcceptanceDate = ((DateTime)item.AcceptanceDate).ToString("yyyy-MM-dd");
                    string AvailableDate = item.AvailableDate == null ? "" : ((DateTime)item.AvailableDate).ToString("yyyyMMdd");
                    string ReCheckDate = item.ReCheckDate == null ? "" : ((DateTime)item.ReCheckDate).ToString("yyyyMMdd");
                    #endregion

                    #region //確認此進貨單是否有進貨檢驗量測單據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.QcRecordId, a.QcStatus
                            FROM MES.QcGoodsReceipt a 
                            WHERE a.GrDetailId = @GrDetailId
                            ORDER BY a.CreateDate DESC";
                    dynamicParameters.Add("GrDetailId", item.GrDetailId);

                    var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                    if (QcGoodsReceiptResult.Count() > 0)
                    {
                        bool qcFlag = false;
                        foreach (var item2 in QcGoodsReceiptResult)
                        {
                            if (item2.QcStatus == "A")
                            {
                                //throw new SystemException("此進貨單身已開立進貨檢驗單據【" + item2.QcRecordId + "】，但尚未量測完成!!");
                            }

                            qcFlag = true;

                            #region //確認此量測單據是否有品異單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.JudgeStatus, a.QcRecordId, a.ReleaseQty
                                    , b.AbnormalqualityNo
                                    FROM QMS.AqBarcode a 
                                    INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                    WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", item2.QcRecordId);

                            var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (AqBarcodeResult.Count() > 0)
                            {
                                foreach (var item3 in AqBarcodeResult)
                                {
                                    if (item3.JudgeStatus == null)
                                    {
                                        throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未進行判定!!");
                                    }
                                    else if (item3.JudgeStatus == "AM")
                                    {
                                        qcFlag = false;
                                    }
                                    else if (item3.JudgeStatus == "R")
                                    {
                                        if (item3.ReleaseQty == null) throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未維護特採數量!!");

                                        #region //確認驗收數量是否正確
                                        if (item.AcceptQty > item3.ReleaseQty) throw new SystemException("驗收數量大於特採數量!!");
                                        #endregion
                                    }
                                    else if (item3.JudgeStatus == "S")
                                    {
                                        #region //若判定為報廢，驗收數量需為0
                                        if (item.AcceptQty != 0) throw new SystemException("此進貨單身判定為報廢，驗收數量需為0!!");
                                        #endregion
                                    }
                                }
                            }
                            else
                            {
                                if (item2.QcStatus != "Y")
                                {
                                    //throw new SystemException("進貨檢驗單號【" + item2.QcRecordId + "】狀態非合格，但查無品異單據!!");
                                }
                            }
                            #endregion
                        }

                        if (qcFlag == false)
                        {
                            //throw new SystemException("此進貨單身已有開立品異單判定為【重新量測】，但尚未有新量測單據!!");
                        }
                    }
                    else
                    {
                        if (item.QcStatus == "1") throw new SystemException("品號【" + item.MtlItemNo + "】為待驗狀態，拋轉前需至少建立一張進貨檢驗單據!!");
                    }
                    #endregion

                    #region //確認是否已有核單者
                    string confirmUserNo = "";
                    if (item.ConfirmUser != null && item.ConfirmUser > 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", item.ConfirmUser);

                        var ConfirmUserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ConfirmUserResult.Count() <= 0)
                        {
                            throw new SystemException($"單身【{item.GrSeq}】: 查詢核單者資料時錯誤!");
                        }

                        foreach (var item2 in ConfirmUserResult)
                        {
                            confirmUserNo = item2.UserNo;
                        }
                    }
                    #endregion

                    #region //INSERT PURTH
                    dynamicParameters = new DynamicParameters();
                    sql = @"INSERT INTO PURTH (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID
                            , TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010, TH011, TH012, TH013, TH014, TH015
                            , TH016, TH017, TH018, TH019, TH020, TH021, TH022, TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030
                            , TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040, TH041, TH042, TH043, TH044, TH045
                            , TH046, TH047, TH048, TH049, TH050, TH051, TH052, TH053, TH054, TH055, TH056, TH057, TH058, TH059, TH060
                            , TH061, TH062, TH063, TH064, TH065, TH066, TH067, TH068, TH069, TH070, TH071, TH072, TH073, TH074, TH500
                            , TH501, TH502, TH503, TH550, TH551, TH552, TH553, TH554, TH555, TH556, TH557, TH558, TH075, TH076, TH077
                            , TH078, TH079, TH080, TH081, TH082, TH083, TH084, TH085, TH086, TH087, TH088, TH089, TH090, TH091, TH092)
                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID
                            , @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010, @TH011, @TH012, @TH013, @TH014, @TH015
                            , @TH016, @TH017, @TH018, @TH019, @TH020, @TH021, @TH022, @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030
                            , @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040, @TH041, @TH042, @TH043, @TH044, @TH045
                            , @TH046, @TH047, @TH048, @TH049, @TH050, @TH051, @TH052, @TH053, @TH054, @TH055, @TH056, @TH057, @TH058, @TH059, @TH060
                            , @TH061, @TH062, @TH063, @TH064, @TH065, @TH066, @TH067, @TH068, @TH069, @TH070, @TH071, @TH072, @TH073, @TH074, @TH500
                            , @TH501, @TH502, @TH503, @TH550, @TH551, @TH552, @TH553, @TH554, @TH555, @TH556, @TH557, @TH558, @TH075, @TH076, @TH077
                            , @TH078, @TH079, @TH080, @TH081, @TH082, @TH083, @TH084, @TH085, @TH086, @TH087, @TH088, @TH089, @TH090, @TH091, @TH092)";
                    dynamicParameters.AddDynamicParams(
                        new
                        {
                            COMPANY = companyNo,
                            CREATOR = UserNo,
                            USR_GROUP,
                            CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                            FLAG = 1,
                            CREATE_TIME = DateTime.Now.ToString("HH:mm:ss"),
                            CREATE_AP = UserNo + "PC",
                            CREATE_PRID = "BM",
                            TH001 = goodsReceipts[0].GrErpPrefix,
                            TH002 = goodsReceipts[0].GrErpNo,
                            TH003 = item.GrSeq,
                            TH004 = item.MtlItemNo,
                            TH005 = item.GrMtlItemName,
                            TH006 = item.GrMtlItemSpec,
                            TH007 = item.ReceiptQty,
                            TH008 = item.UomNo,
                            TH009 = item.InventoryNo,
                            TH010 = item.LotNumber,
                            TH011 = item.PoErpPrefix,
                            TH012 = item.PoErpNo,
                            TH013 = item.PoSeq,
                            TH014 = item.ErpAcceptanceDate,
                            TH015 = item.AcceptQty,
                            TH016 = item.AvailableQty,
                            TH017 = item.ReturnQty,
                            TH018 = item.OrigUnitPrice,
                            TH019 = item.OrigAmount,
                            TH020 = item.OrigDiscountAmt,
                            TH021 = item.TrErpPrefix,
                            TH022 = item.TrErpNo,
                            TH023 = item.TrSeq,
                            TH024 = item.ReceiptExpense,
                            TH025 = item.DiscountDescription,
                            TH026 = item.PaymentHold,
                            TH027 = item.Overdue,
                            TH028 = item.QcStatus,
                            TH029 = item.ReturnCode,
                            TH030 = item.ConfirmStatus,
                            TH031 = item.CloseStatus,
                            TH032 = item.ReNewStatus,
                            TH033 = item.Remark,
                            TH034 = item.InventoryQty,
                            TH035 = item.SmallUnit,
                            TH036 = AvailableDate,
                            TH037 = ReCheckDate,
                            TH038 = confirmUserNo,
                            TH039 = item.ApvErpPrefix,
                            TH040 = item.ApvErpNo,
                            TH041 = item.ApvSeq,
                            TH042 = item.ProjectCode,
                            TH043 = item.ExpenseEntry,
                            TH044 = item.PremiumAmountFlag,
                            TH045 = item.OrigPreTaxAmt,
                            TH046 = item.OrigTaxAmt,
                            TH047 = item.PreTaxAmt,
                            TH048 = item.TaxAmt,
                            TH049 = item.ReceiptPackageQty,
                            TH050 = item.AcceptancePackageQty,
                            TH051 = item.ReturnPackageQty,
                            TH052 = item.PremiumAmount,
                            TH053 = item.PackageUnit,
                            TH054 = item.ReserveTaxCode,
                            TH055 = item.OrigPremiumAmount,
                            TH056 = item.AvailableUomNo,
                            TH057 = item.OrigCustomer,
                            TH058 = item.ApproveStatus,
                            TH059 = item.EbcErpPreNo,
                            TH060 = item.EbcEdition,
                            TH061 = item.ProductSeqAmount,
                            TH062 = item.MtlItemType,
                            TH063 = item.Loaction,
                            TH064 = "",
                            TH065 = "",
                            TH066 = "",
                            TH067 = 0.000000,
                            TH068 = 0.000000,
                            TH069 = "",
                            TH070 = "",
                            TH071 = "",
                            TH072 = "",
                            TH073 = item.TaxRate,
                            TH074 = item.MultipleFlag,
                            TH500 = "",
                            TH501 = "",
                            TH502 = "",
                            TH503 = "",
                            TH550 = "",
                            TH551 = "",
                            TH552 = "",
                            TH553 = "",
                            TH554 = "",
                            TH555 = "",
                            TH556 = "",
                            TH557 = "",
                            TH558 = "",
                            TH075 = "",
                            TH076 = "",
                            TH077 = "",
                            TH078 = "",
                            TH079 = "",
                            TH080 = "",
                            TH081 = "",
                            TH082 = item.GrErpStatus,
                            TH083 = "",
                            TH084 = "",
                            TH085 = "",
                            TH086 = "",
                            TH087 = "",
                            TH088 = item.TaxCode,
                            TH089 = item.DiscountRate,
                            TH090 = item.DiscountPrice,
                            TH091 = item.QcType,
                            TH092 = ""
                        });
                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion
                }
                #endregion

                #region //UPDATE SCM.GoodsReceipt
                dynamicParameters = new DynamicParameters();
                sql = @"UPDATE SCM.GoodsReceipt SET
                        GrErpNo = @GrErpNo,
                        TransferStatus = 'Y',
                        TransferDate = @TransferDate,
                        LastModifiedDate = @LastModifiedDate,
                        LastModifiedBy = @LastModifiedBy
                        WHERE GrId = @GrId";
                dynamicParameters.AddDynamicParams(
                  new
                  {
                      goodsReceipts[0].GrErpNo,
                      TransferDate = DateTime.Now,
                      LastModifiedDate,
                      LastModifiedBy = UserId,
                      GrId
                  });

                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                #endregion

                #region //UPDATE SCM.GrDetail
                dynamicParameters = new DynamicParameters();
                sql = @"UPDATE SCM.GrDetail SET
                        GrErpNo = @GrErpNo,
                        TransferStatus = 'Y',
                        TransferDate = @TransferDate,
                        LastModifiedDate = @LastModifiedDate,
                        LastModifiedBy = @LastModifiedBy
                        WHERE GrId = @GrId";
                dynamicParameters.AddDynamicParams(
                  new
                  {
                      goodsReceipts[0].GrErpNo,
                      TransferDate = DateTime.Now,
                      LastModifiedDate,
                      LastModifiedBy = UserId,
                      GrId
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateGrTotal -- 更新進貨單頭統整資料 -- Ann 2024-03-07
        public string UpdateGrTotal(int GrId, double? Exchange, SqlConnection sqlConnection, SqlConnection sqlConnection2, int UserId)
        {
            try
            {
                int rowsAffected = 0;
                int OriUnitDecimal = 0; //原幣單價取位
                int OriPriceDecimal = 0; //原幣金額取位
                int UnitDecimal = 0; //本幣單價取位
                int PriceDecimal = 0; //本幣金額取位

                double TotalAmount = 0;
                double DeductAmount = 0;
                double OrigTax = 0;
                double ReceiptAmount = 0;
                double Quantity = 0;
                double OrigPreTaxAmount = 0;
                double PreTaxAmount = 0;
                double TaxAmount = 0;

                #region //確認進貨單頭資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.CurrencyCode, a.Exchange, a.TaxRate, a.TaxType
                        FROM SCM.GoodsReceipt a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                if (GoodsReceiptResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");

                string OriCurrencyCode = "";
                double OriExchange = -1;
                double TaxRate = -1;
                string TaxType = "";
                foreach (var item in GoodsReceiptResult)
                {
                    OriCurrencyCode = item.CurrencyCode;
                    OriExchange = item.Exchange;
                    TaxRate = item.TaxRate;
                    TaxType = item.TaxType;
                    if (Exchange < 0) Exchange = OriExchange;
                }
                #endregion

                #region //取得進貨單身資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.GrDetailId, a.AvailableQty, a.OrigUnitPrice, a.OrigDiscountAmt, a.OrigAmount
                        FROM SCM.GrDetail a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);
                #endregion

                #region //原幣小數點取位
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.MF003, a.MF004
                        FROM CMSMF a 
                        WHERE a.MF001 = @Currency";
                dynamicParameters.Add("Currency", OriCurrencyCode);

                var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                foreach (var item in CMSMFResult)
                {
                    OriUnitDecimal = Convert.ToInt32(item.MF003);
                    OriPriceDecimal = Convert.ToInt32(item.MF004);
                }
                #endregion

                #region //取得本幣幣別
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) CurrencyCode
                        FROM CMSMA
                        WHERE 1=1";

                var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                string CurrencyCode = "";
                foreach (var item in CMSMAResult)
                {
                    CurrencyCode = item.CurrencyCode;
                }
                #endregion

                #region //本幣小數點取位
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.MF003, a.MF004
                        FROM CMSMF a 
                        WHERE a.MF001 = @Currency";
                dynamicParameters.Add("Currency", CurrencyCode);

                var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                foreach (var item in CMSMFResult2)
                {
                    UnitDecimal = Convert.ToInt32(item.MF003);
                    PriceDecimal = Convert.ToInt32(item.MF004);
                }
                #endregion

                #region //更新單身金額
                foreach (var item in GrDetailResult)
                {
                    JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(item.AvailableQty, item.OrigUnitPrice, item.OrigDiscountAmt, item.OrigAmount, TaxRate, TaxType, Exchange
                        , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                    if (calculateTaxAmtResult["status"].ToString() == "success")
                    {
                        double OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                        double OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                        double PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                        double TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                        double OrigAmount = Convert.ToDouble(calculateTaxAmtResult["OrigAmount"]);

                        #region //Update SCM.GrDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                OrigPreTaxAmt = @OrigPreTaxAmt,
                                OrigTaxAmt = @OrigTaxAmt,
                                PreTaxAmt = @PreTaxAmt,
                                TaxAmt = @TaxAmt,
                                OrigAmount = @OrigAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OrigPreTaxAmt,
                                OrigTaxAmt,
                                PreTaxAmt,
                                TaxAmt,
                                OrigAmount,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                item.GrDetailId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }
                    else
                    {
                        throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                    }
                }
                #endregion

                dynamicParameters = new DynamicParameters();
                sql = @"SELECT SUM(a.OrigAmount) TotalAmount, SUM(a.OrigDiscountAmt) DeductAmount
                        , SUM(a.OrigTaxAmt) OrigTax, SUM(a.ReceiptExpense) ReceiptAmount
                        , SUM(a.AvailableQty) Quantity, SUM(a.OrigPreTaxAmt) OrigPreTaxAmount
                        , SUM(a.PreTaxAmt) PreTaxAmount, SUM(a.TaxAmt) TaxAmount
                        FROM SCM.GrDetail a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GrDetailQtyResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in GrDetailQtyResult)
                {
                    TotalAmount = Convert.ToDouble(item.TotalAmount);
                    DeductAmount = Convert.ToDouble(item.DeductAmount);
                    OrigTax = Convert.ToDouble(item.OrigTax);
                    ReceiptAmount = Convert.ToDouble(item.ReceiptAmount);
                    Quantity = Convert.ToDouble(item.Quantity);
                    OrigPreTaxAmount = Convert.ToDouble(item.OrigPreTaxAmount);
                    PreTaxAmount = Convert.ToDouble(item.PreTaxAmount);
                    TaxAmount = Convert.ToDouble(item.TaxAmount);
                }

                #region //Update SCM.GoodsReceipt
                dynamicParameters = new DynamicParameters();
                sql = @"UPDATE SCM.GoodsReceipt SET
                        Exchange = @Exchange,
                        TotalAmount = @TotalAmount,
                        DeductAmount = @DeductAmount,
                        OrigTax = @OrigTax,
                        ReceiptAmount = @ReceiptAmount,
                        Quantity = @Quantity,
                        OrigPreTaxAmount = @OrigPreTaxAmount,
                        PreTaxAmount = @PreTaxAmount,
                        TaxAmount = @TaxAmount,
                        LastModifiedDate = @LastModifiedDate,
                        LastModifiedBy = @LastModifiedBy
                        WHERE GrId = @GrId";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        Exchange,
                        TotalAmount = Math.Round(TotalAmount, OriPriceDecimal),
                        DeductAmount = Math.Round(DeductAmount, OriPriceDecimal),
                        OrigTax = Math.Round(OrigTax, OriPriceDecimal),
                        ReceiptAmount = Math.Round(ReceiptAmount, PriceDecimal),
                        Quantity,
                        OrigPreTaxAmount = Math.Round(OrigPreTaxAmount, OriPriceDecimal),
                        PreTaxAmount = Math.Round(PreTaxAmount, PriceDecimal),
                        TaxAmount = Math.Round(TaxAmount, PriceDecimal),
                        LastModifiedDate,
                        LastModifiedBy = UserId,
                        GrId
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //CalculateTaxAmt -- 計算稅額 -- Ann 2024-03-14
        public string CalculateTaxAmt(double? AvailableQty, double? OrigUnitPrice, double? OrigDiscountAmt, double? OrigAmount, double? TaxRate, string TaxType, double? Exchange
            , int OriUnitDecimal, int OriPriceDecimal, int UnitDecimal, int PriceDecimal)
        {
            try
            {
                if (OriUnitDecimal < 0) throw new SystemException("【原幣單價小數點進位數】格式錯誤!");
                if (OriPriceDecimal < 0) throw new SystemException("【原幣金額小數點進位數】格式錯誤!");
                if (UnitDecimal < 0) throw new SystemException("【本幣單價小數點進位數】格式錯誤!");
                if (PriceDecimal < 0) throw new SystemException("【本幣金額小數點進位數】格式錯誤!");
                if (TaxType.Length <= 0) throw new SystemException("【課稅別】格式錯誤!");

                double OrigPreTaxAmt = 0; //原幣未稅金額
                double OrigTaxAmt = 0; //原幣稅金額
                double PreTaxAmt = 0; //本幣未稅金額
                double TaxAmt = 0; //本幣稅金額
                OrigAmount = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice), OriPriceDecimal);

                switch (TaxType)
                {
                    #region //應稅內含
                    case "1":
                        OrigPreTaxAmt = Math.Round(Convert.ToDouble((AvailableQty * OrigUnitPrice - OrigDiscountAmt) / (1 + TaxRate)), OriPriceDecimal);
                        OrigTaxAmt = Math.Round(Convert.ToDouble(OrigAmount - OrigPreTaxAmt), OriPriceDecimal);
                        PreTaxAmt = Math.Round(Convert.ToDouble(OrigPreTaxAmt * Exchange), PriceDecimal);
                        TaxAmt = Math.Round(Convert.ToDouble(OrigTaxAmt * Exchange), PriceDecimal);
                        break;
                    #endregion
                    #region //應稅外加
                    case "2":
                        OrigPreTaxAmt = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice - OrigDiscountAmt), OriPriceDecimal);
                        OrigTaxAmt = Math.Round(Convert.ToDouble((AvailableQty * OrigUnitPrice - OrigDiscountAmt) * TaxRate), OriPriceDecimal);
                        PreTaxAmt = Math.Round(Convert.ToDouble(OrigPreTaxAmt * Exchange), PriceDecimal);
                        TaxAmt = Math.Round(Convert.ToDouble(OrigTaxAmt * Exchange), PriceDecimal);
                        break;
                    #endregion
                    #region //免稅
                    case "3":
                    case "4":
                    case "9":
                        OrigPreTaxAmt = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice - OrigDiscountAmt), OriPriceDecimal);
                        OrigTaxAmt = 0;
                        PreTaxAmt = Math.Round(Convert.ToDouble(OrigPreTaxAmt * Exchange), PriceDecimal);
                        TaxAmt = 0;
                        break;
                    #endregion
                    #region //例外狀況
                    default:
                        throw new SystemException("稅別碼資料錯誤!!");
                        #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    OrigPreTaxAmt,
                    OrigTaxAmt,
                    PreTaxAmt,
                    TaxAmt,
                    OrigAmount
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
            }

            return jsonResponse.ToString();
        }
        #endregion

        public string ConfirmGrDetail(int GrdetailId, SqlConnection sqlConnection, SqlConnection sqlConnection2, int UserId)
        {
            try
            {
                int rowsAffected = 0;
                string grseq = "";
                string grerpno = "";

                var ErpNo = "";
                string USR_GROUP = "";
                int CurrentCompany = 0;
                string ReceiptDate = "";

                string GrErpPrefix = "";
                string GrErpNo = "";

                string MaxReceiptDate = "";

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");
                string dateNowMonth = DateTime.Now.ToString("yyyyMM");

                #region //參數
                string UserNo = "";
                int CompanyIdBase = -1;
                int grId = -1;
                var TH001 = "";
                var TH002 = "";
                var TH003 = "";
                var TH004 = "";
                var TH005 = "";
                var TH006 = "";
                var TH007 = 0.000000;
                var TH008 = "";
                var TH009 = "";
                var TH010 = "";
                var TH011 = "";
                var TH012 = "";
                var TH013 = "";
                var TH014 = "";
                var TH015 = 0.000000;
                var TH016 = 0.000000;
                var TH017 = 0.000000;
                var TH018 = 0.000000;
                var TH019 = 0.000000;
                var TH020 = 0.000000;
                var TH021 = "";
                var TH022 = "";
                var TH023 = "";
                var TH024 = 0.000000;
                var TH025 = "";
                var TH026 = "";
                var TH027 = "";
                var TH028 = "";
                var TH029 = "";
                var TH030 = "";
                var TH031 = "";
                var TH032 = "";
                var TH033 = "";
                var TH034 = 0.000000;
                var TH035 = "";
                var TH036 = "";
                var TH037 = "";
                var TH038 = "";
                var TH039 = "";
                var TH040 = "";
                var TH041 = "";
                var TH042 = "";
                var TH043 = "";
                var TH044 = "";
                var TH045 = 0.000000;
                var TH046 = 0.000000;
                var TH047 = 0.000000;
                var TH048 = 0.000000;
                var TH049 = 0.000000;
                var TH050 = 0.000000;
                var TH051 = 0.000000;
                var TH052 = 0.000000;
                var TH053 = "";
                var TH054 = "";
                var TH055 = 0.000000;
                var TH056 = "";
                var TH057 = "";
                var TH058 = "";
                var TH059 = "";
                var TH060 = "";
                var TH061 = 0.000000;
                var TH062 = "";
                var TH063 = "";
                var TH064 = "";
                var TH065 = "";
                var TH066 = "";
                var TH067 = 0.000000;
                var TH068 = 0.000000;
                var TH069 = "";
                var TH070 = "";
                var TH071 = "";
                var TH072 = "";
                var TH073 = 0.000000;
                var TH074 = "";
                var TH500 = "";
                var TH501 = "";
                var TH502 = "";
                var TH503 = "";
                var TH550 = "";
                var TH551 = "";
                var TH552 = "";
                var TH553 = "";
                var TH554 = "";
                var TH555 = "";
                var TH556 = "";
                var TH557 = "";
                var TH558 = "";
                var TH075 = "";
                var TH076 = "";
                var TH077 = "";
                var TH078 = "";
                var TH079 = "";
                var TH080 = "";
                var TH081 = "";
                var TH082 = "";
                var TH083 = "";
                var TH084 = "";
                var TH085 = "";
                var TH086 = "";
                var TH087 = "";
                var TH088 = "";
                var TH089 = 0.000000;
                var TH090 = 0.000000;
                var TH091 = "";
                var TH092 = "";

                string CloseStatus = "";
                string PoErpPrefix = ""; //TH011 採購單別 單號 序號
                string PoErpNo = ""; //TH012 
                string PoSeq = ""; //TH013 
                #endregion

                #region //取得目前使用者資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.UserNo, b.CompanyId
                                    FROM BAS.[User] a
									INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                    WHERE UserId = @UserId";
                dynamicParameters.Add("UserId", UserId);

                var UserResult = sqlConnection.Query(sql, dynamicParameters);

                //string UserNo = "";
                foreach (var item in UserResult)
                {
                    UserNo = item.UserNo;
                    CurrentCompany = item.CompanyId;
                }
                #endregion

                #region //確認公司別DB
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                dynamicParameters.Add("CompanyId", CurrentCompany);

                var resultCompany = sqlConnection.Query(sql, dynamicParameters);

                if (resultCompany.Count() <= 0) throw new SystemException("公司別錯誤!");

                foreach (var item in resultCompany)
                {

                    ErpNo = item.ErpNo;
                }
                #endregion

                #region //找該單身
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT b.GrDetailId, b.GrId, b.GrErpPrefix, b.GrErpNo, b.GrSeq, b.MtlItemId, b.MtlItemNo, b.GrMtlItemName, b.GrMtlItemSpec
                                , b.ReceiptQty, b.UomId, b.UomNo, b.InventoryId, b.InventoryNo, b.LotNumber, b.PoErpPrefix, b.PoErpNo, b.PoSeq
                                , b.AcceptanceDate, CASE WHEN LEN(b.AcceptanceDate) > 0 THEN FORMAT(CAST(b.AcceptanceDate as date), 'yyyyMMdd') ELSE '' END AS ErpAcceptanceDate
                                , b.AcceptQty, b.AvailableQty, b.ReturnQty, b.OrigUnitPrice, b.OrigAmount, b.OrigDiscountAmt, b.TrErpPrefix, b.TrErpNo, b.TrSeq
                                , b.ReceiptExpense, b.DiscountDescription, b.PaymentHold, b.Overdue, b.QcStatus, b.ReturnCode, b.ConfirmStatus, b.CloseStatus
                                , b.ReNewStatus, b.Remark, b.InventoryQty, b.SmallUnit
                                , CASE WHEN FORMAT(b.AvailableDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(b.AvailableDate as date), 'yyyy-MM-dd') END AS AvailableDate
                                , CASE WHEN LEN(b.AvailableDate) > 0 THEN FORMAT(CAST(b.AvailableDate as date), 'yyyyMMdd') ELSE '' END AS ErpAvailableDate
                                , CASE WHEN FORMAT(b.ReCheckDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(b.ReCheckDate as date), 'yyyy-MM-dd') END AS ReCheckDate
                                , CASE WHEN LEN(b.ReCheckDate) > 0 THEN FORMAT(CAST(b.ReCheckDate as date), 'yyyyMMdd') ELSE '' END AS ErpReCheckDate
                                , b.ConfirmUser, b.ApvErpPrefix, b.ApvErpNo
                                , b.ApvSeq, b.ProjectCode, b.ExpenseEntry, b.PremiumAmountFlag, b.OrigPreTaxAmt, b.OrigTaxAmt, b.PreTaxAmt, b.TaxAmt, b.ReceiptPackageQty
                                , b.AcceptancePackageQty, b.ReturnPackageQty, b.PremiumAmount, b.PackageUnit, b.ReserveTaxCode, b.OrigPremiumAmount, b.AvailableUomId
                                , b.AvailableUomNo, b.OrigCustomer, b.ApproveStatus, b.EbcErpPreNo, b.EbcEdition, b.ProductSeqAmount, b.MtlItemType, b.Loaction
                                , b.TaxRate, b.MultipleFlag, b.GrErpStatus, b.TaxCode, b.DiscountRate, b.DiscountPrice, b.QcType, b.TransferStatus, b.TransferDate
                                , b.CreateDate, b.LastModifiedDate, b.CreateBy, b.LastModifiedBy
                                , a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate,a.TransferStatus
								,x.AcceptanceDateCount,a.GrId
                                FROM SCM.GoodsReceipt a
                                INNER JOIN SCM.GrDetail b on a.GrId = b.GrId
                                OUTER APPLY(
                                    SELECT CASE WHEN COUNT(DISTINCT FORMAT(x1.AcceptanceDate, 'yyyyMM'))  >1 
                                    THEN 'N' ELSE 'Y' 
                                    END AS AcceptanceDateCount
                                    FROM SCM.GrDetail x1
                                    WHERE x1.GrId = a.GrId
                                ) x
                                WHERE b.GrDetailId = @GrDetailId";
                dynamicParameters.Add("GrDetailId", GrdetailId);

                var GRDetailResult = sqlConnection.Query(sql, dynamicParameters);

                if (GRDetailResult.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");

                foreach (var item in GRDetailResult)
                {
                    grseq = item.GrSeq;
                    grerpno = item.GrErpFullNo;
                    ReceiptDate = item.ReceiptDate;
                    GrErpPrefix = item.GrErpPrefix;
                    GrErpNo = item.GrErpNo;

                    #region //參數
                    CompanyIdBase = item.CompanyId;
                    GrErpPrefix = item.GrErpPrefix;
                    GrErpNo = item.GrErpNo;
                    ReceiptDate = item.ReceiptDate;
                    string AvailableDate = item.AvailableDate == null ? "" : ((DateTime)item.AvailableDate).ToString("yyyyMMdd");
                    string ReCheckDate = item.ReCheckDate == null ? "" : ((DateTime)item.ReCheckDate).ToString("yyyyMMdd");

                    TH001 = item.GrErpPrefix;
                    TH002 = item.GrErpNo;
                    TH003 = item.GrSeq;
                    TH004 = item.MtlItemNo;
                    TH005 = item.GrMtlItemName;
                    TH006 = item.GrMtlItemSpec;
                    TH007 = item.ReceiptQty;
                    TH008 = item.UomNo;
                    TH009 = item.InventoryNo;
                    TH010 = item.LotNumber;
                    TH011 = item.PoErpPrefix;
                    TH012 = item.PoErpNo;
                    TH013 = item.PoSeq;
                    TH014 = item.ErpAcceptanceDate != null ? item.ErpAcceptanceDate : "";
                    TH015 = item.AcceptQty;
                    TH016 = item.AvailableQty;
                    TH017 = item.ReturnQty;
                    TH018 = item.OrigUnitPrice;
                    TH019 = item.OrigAmount;
                    TH020 = item.OrigDiscountAmt;
                    TH021 = item.TrErpPrefix;
                    TH022 = item.TrErpNo;
                    TH023 = item.TrSeq;
                    TH024 = item.ReceiptExpense;
                    TH025 = item.DiscountDescription;
                    TH026 = item.PaymentHold;
                    TH027 = item.Overdue;
                    TH028 = item.QcStatus;
                    TH029 = item.ReturnCode;
                    TH030 = item.ConfirmStatus;
                    TH031 = item.CloseStatus;
                    TH032 = item.ReNewStatus;
                    TH033 = item.Remark;
                    TH034 = item.InventoryQty;
                    TH035 = item.SmallUnit;
                    TH036 = AvailableDate;
                    TH037 = ReCheckDate;
                    //TH038 = item.ConfirmUser;
                    TH039 = item.ApvErpPrefix;
                    TH040 = item.ApvErpNo;
                    TH041 = item.ApvSeq;
                    TH042 = item.ProjectCode;
                    TH043 = item.ExpenseEntry;
                    TH044 = item.PremiumAmountFlag;
                    TH045 = item.OrigPreTaxAmt;
                    TH046 = item.OrigTaxAmt;
                    TH047 = item.PreTaxAmt;
                    TH048 = item.TaxAmt;
                    TH049 = item.ReceiptPackageQty;
                    TH050 = item.AcceptancePackageQty;
                    TH051 = item.ReturnPackageQty;
                    TH052 = item.PremiumAmount;
                    TH053 = item.PackageUnit;
                    TH054 = item.ReserveTaxCode;
                    TH055 = item.OrigPremiumAmount;
                    TH056 = item.AvailableUomNo;
                    TH057 = item.OrigCustomer;
                    TH058 = item.ApproveStatus;
                    TH059 = item.EbcErpPreNo;
                    TH060 = item.EbcEdition;
                    TH061 = item.ProductSeqAmount;
                    TH062 = item.MtlItemType;
                    TH063 = item.Loaction;
                    TH064 = "";
                    TH065 = "";
                    TH066 = "";
                    TH067 = 0.000000;
                    TH068 = 0.000000;
                    TH069 = "";
                    TH070 = "";
                    TH071 = "";
                    TH072 = "";
                    TH073 = item.TaxRate;
                    TH074 = item.MultipleFlag;
                    TH500 = "";
                    TH501 = "";
                    TH502 = "";
                    TH503 = "";
                    TH550 = "";
                    TH551 = "";
                    TH552 = "";
                    TH553 = "";
                    TH554 = "";
                    TH555 = "";
                    TH556 = "";
                    TH557 = "";
                    TH558 = "";
                    TH075 = "";
                    TH076 = "";
                    TH077 = "";
                    TH078 = "";
                    TH079 = "";
                    TH080 = "";
                    TH081 = "";
                    TH082 = item.GrErpStatus;
                    TH083 = "";
                    TH084 = "";
                    TH085 = "";
                    TH086 = "";
                    TH087 = "";
                    TH088 = item.TaxCode;
                    TH089 = item.DiscountRate;
                    TH090 = item.DiscountPrice;
                    TH091 = item.QcType;
                    TH092 = "";
                    //GrDetailId = item.GrDetailId;
                    grId = item.GrId;
                    #endregion



                    #region //判斷進貨檢驗單據條件
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.QcRecordId, a.QcStatus
                                    , b.GrSeq
                                    FROM MES.QcGoodsReceipt a 
                                    INNER JOIN SCM.GrDetail b ON a.GrDetailId = b.GrDetailId
                                    WHERE a.GrDetailId = @GrDetailId
                                    ORDER BY a.CreateDate DESC";
                    dynamicParameters.Add("GrDetailId", GrdetailId);

                    var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                    if (QcGoodsReceiptResult.Count() > 0)
                    {
                        bool qcFlag = false;
                        foreach (var item2 in QcGoodsReceiptResult)
                        {

                            if (item2.QcStatus == "A")
                            {
                                //throw new SystemException("單身序號【" + grseq + "】已開立進貨檢驗單據【" + item2.QcRecordId + "】，但尚未量測完成!!");
                            }

                            qcFlag = true;

                            #region //確認此量測單據是否有品異單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.JudgeStatus, a.QcRecordId, a.ReleaseQty
                                                , b.AbnormalqualityNo
                                                FROM QMS.AqBarcode a 
                                                INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                                WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", item2.QcRecordId);

                            var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (AqBarcodeResult.Count() > 0)
                            {
                                foreach (var item3 in AqBarcodeResult)
                                {
                                    if (item3.JudgeStatus == null)
                                    {
                                        throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未進行判定!!");
                                    }
                                    else if (item3.JudgeStatus == "AM")
                                    {
                                        qcFlag = false;
                                    }
                                    else if (item3.JudgeStatus == "R")
                                    {
                                        if (item3.ReleaseQty == null) throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未維護特採數量!!");

                                        #region //確認驗收數量是否正確
                                        if (item.AcceptQty > item3.ReleaseQty) throw new SystemException("驗收數量大於特採數量!!");
                                        #endregion
                                    }
                                    else if (item3.JudgeStatus == "S")
                                    {
                                        #region //若判定為報廢，驗收數量需為0
                                        if (item.AcceptQty != 0) throw new SystemException("單身序號【" + item2.GrSeq + "】判定為報廢，驗收數量需為0!!");
                                        #endregion
                                    }
                                }
                            }
                            else
                            {
                                if (item2.QcStatus != "Y")
                                {
                                    //throw new SystemException("進貨檢驗單號【" + item2.QcRecordId + "】狀態非合格，但查無品異單據!!");
                                }
                            }
                            #endregion
                        }

                        if (qcFlag == false)
                        {
                            //throw new SystemException("單身序號【" + grseq + "】已有開立品異單判定為【重新量測】，但尚未有新量測單據!!");
                        }
                    }
                    else
                    {
                        //if (item.QcStatus == "1") throw new SystemException("品號【" + item.MtlItemNo + "】為待驗狀態，需進行檢驗!!");
                    }
                    #endregion

                    

                    #region 



                    #endregion

                }

                #endregion


                #region //ERP

                #region //審核ERP權限
                USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "PURI13", "CONFIRM");
                #endregion

                #region //判斷開立日期是否超過結帳日
                string CloseDateBase = "";
                string ReceiveDateBase = ReceiptDate;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT MA011,MA012,MA013
                                FROM CMSMA ";
                var resultCloseDate = sqlConnection2.Query(sql, dynamicParameters);
                foreach (var item in resultCloseDate)
                {
                    CloseDateBase = item.MA013;
                }

                DateTime docDateBase;
                DateTime closeDateBase;

                if (DateTime.TryParseExact(ReceiveDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                    DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                {
                    int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                    if (resultDate <= 0)
                    {
                        throw new SystemException("【進貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                    }
                }
                else
                {
                    throw new SystemException("日期字符串无效，无法比较");
                }
                #endregion

                #region //取出進貨單單據資料 + 異動更新相關資料表
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.TG003,a.TG005,a.TG007,a.TG013,a.TG014
                                ,b.TH003,b.TH004,b.TH007,b.TH009,b.TH010
                                ,b.TH011,b.TH012,b.TH013
                                ,b.TH014,b.TH015,b.TH016,b.TH017,b.TH030,b.TH031,b.TH063
                                ,b.TH021,b.TH022,b.TH023
                                ,b.TH036,b.TH037,b.TH047
                                ,c.MB022
                                ,d.MQ019,md3.MF005
                                ,ISNULL(md1.GrRate,1) GrRate
                                ,ISNULL(md2.GrAvailableRate,1) GrAvailableRate
                                ,(b.TH015*ISNULL(md1.GrRate,1)) ConversionQty
                                ,(b.TH016*ISNULL(md2.GrAvailableRate,1)) ConversionAvailableQty
                                ,CASE 
                                    WHEN b.TH016 != 0 THEN ROUND((b.TH047 / (b.TH016*ISNULL(md2.GrAvailableRate,1))),CAST(md3.MF005 AS INT)) 
                                    ELSE 0
                                END AS UnitCost
                                FROM PURTG a
                                INNER JOIN PURTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                INNER JOIN CMSMQ d ON d.MQ001=a.TG001 AND d.MQ003='34'
                                INNER JOIN CMSMC e ON e.MC001=b.TH009
                                LEFT JOIN INVMD f ON f.MD001=b.TH004
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH056
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE b.TH001 = @GrErpPrefix
                                AND b.TH002 = @GrErpNo
                                AND b.TH003 = @GrSeq";
                dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                dynamicParameters.Add("GrErpNo", GrErpNo);
                dynamicParameters.Add("GrSeq", grseq);

                var resultRoErp = sqlConnection2.Query(sql, dynamicParameters);
                if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP進貨單單據!");
                foreach (var item in resultRoErp)
                {
                    if (item.TG013 == "Y") throw new SystemException("該進貨單單據已經核單,不能操作");
                    if (item.TG013 == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                    if (item.TH030 == "Y") throw new SystemException("該進貨單單據單身已經核單,不能操作");
                    if (item.TH031 == "Y") throw new SystemException("該進貨單單據已經結帳,不能操作");
                    if (item.TH021 != "") throw new SystemException("來源暫入單模式尚未開發,不能操作");
                    if (item.MQ019 == "N") throw new SystemException("不核對採購單模式尚未開發,不能操作");

                    #region //欄位參數撈取
                    string ReceiveDateErp = item.TG003; //TG003 進貨日期
                    string SupplierNo = item.TG005; //TG005 供應商
                    string CurrencyCode = item.TG007; //TG007 幣別
                    Double Exchange = Convert.ToDouble(item.TG008); //TG008 匯率
                    string ConfirmStatus = item.TG013; //TG013 單頭確認
                    string DocDate = item.TG014; //TG013 單頭確認

                    //string GrSeq = item.TH003; //TH003 進貨單序號
                    string GrFull = GrErpPrefix + '-' + GrErpNo + '-' + grseq;
                    string MtlItemNo = item.TH004; //TH004 品號
                    string InventoryNo = item.TH009; //TH009 庫別
                    string LotNumber = item.TH010; //TH010 批號
                    PoErpPrefix = item.TH011; //TH011 採購單別 單號 序號
                    PoErpNo = item.TH012; //TH012 
                    PoSeq = item.TH013; //TH013 
                    string PoFull = item.TH011 + '-' + item.TH012 + '-' + item.TH013;
                    //string TH014 = item.TH014; //TH014 驗收日期
                    Double ReceiptQty = Convert.ToDouble(item.TH007); //TH014 進貨數量
                    Double AcceptQty = Convert.ToDouble(item.TH015); //TH014 驗收數量
                    Double AvailableQty = Convert.ToDouble(item.TH016); //TH014 計價數量 
                    Double ReturnQty = Convert.ToDouble(item.TH017); //TH014 驗退數量
                                                                     //string TH030 = item.TH030; //TH030 確認碼
                    CloseStatus = item.TH031; //TH031 結帳碼
                    string AvailableDate = item.TH036; //TH036 有效日期
                    string ReCheckDate = item.TH037; //TH037 複檢日期
                    string Location = item.TH063; //TH060 儲位
                    string TrErpPrefix = item.TH021; //TH021 暫入單別 / TH022 暫入單號 / TH023 暫入序號
                    string TrErpNo = item.TH022;
                    string TrSeq = item.TH023;
                    string TsnFull = item.TH021 + '-' + item.TH022 + '-' + item.TH023;
                    string LotManagement = item.MB022; //品號批號管理

                    Double GrRate = Convert.ToDouble(item.GrRate); //單位換算率
                    Double ConversionQty = Convert.ToDouble(TH015 * item.GrRate); //單位換算後數量
                    Double ConversionAvailableQty = Convert.ToDouble(TH016 * item.GrAvailableRate); //單位換算後計價數量
                    Double UnitCost = Convert.ToDouble(TH047 / (TH016 * item.GrAvailableRate)); //單位成本
                                                                                                //Double TH047 = Convert.ToDouble(item.TH047); //本幣未稅金額
                                                                                                // Double TH018 = Convert.ToDouble(item.TH018); //原幣單價
                    Double PoRate = 0; //採購單位換算率
                    Double PoAvailableRate = 0; //採購計價單位換算率
                    Double PoConversionQty = 0; //採購單位換算後數量
                    Double PoConversionAvailableQty = 0; //採購單位換算後計價數量

                    DateTime dateC = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                    DateTime dateD = DateTime.ParseExact(TH014, "yyyyMMdd", null);
                    if (dateC > dateD)
                    {
                        throw new SystemException("資料異常:單身驗收日不可以小於單據日!!");
                    }
                    if (MaxReceiptDate != "")
                    {
                        DateTime dateA = DateTime.ParseExact(item.TH014, "yyyyMMdd", null);
                        DateTime dateB = DateTime.ParseExact(MaxReceiptDate, "yyyyMMdd", null);
                        //MaxReceiptDate = date.ToString("yyyy-MM-dd"); // yyyy-MM-dd 格式


                        //DateTime dateA = DateTime.Parse(item.TH014);
                        //DateTime dateB = DateTime.Parse(MaxReceiptDate);
                        if (dateA > dateB)
                        {
                            MaxReceiptDate = item.TH014;
                        }
                    }
                    else
                    {
                        MaxReceiptDate = item.TH014;
                    }
                    #endregion

                    #region //檢查段

                    #region //數量查驗
                    if (ReceiptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】進貨數量不可小於0!!!");
                    if (AcceptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可小於0!!!");
                    if (AvailableQty < 0) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可小於0!!!");
                    if (ReceiptQty < AcceptQty) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可大於進貨數量!!!");
                    if (AcceptQty < AvailableQty) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可大於驗收數量!!!");
                    #endregion

                    #region //PURTD 採購單
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 ISNULL(a.TD008,0) TD008,ISNULL(a.TD015,0) TD015,a.TD016,a.TD018,
                                    ISNULL(md1.PoRate,1) PoRate, 
                                    ISNULL(md2.PoAvailableRate,1) PoAvailableRate, 
                                    (a.TD008*ISNULL(md1.PoRate,1)) PoConversionQty,
                                    (a.TD015*ISNULL(md1.PoRate,1)) PoConversionSiQty,
                                    (a.TD058*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableQty,
                                    (a.TD060*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableSiQty
                                    FROM PURTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD059
                                    ) md2
                                    WHERE a.TD001 = @TD001
                                    AND a.TD002 = @TD002
                                    AND a.TD003 = @TD003";
                    dynamicParameters.Add("TD001", PoErpPrefix);
                    dynamicParameters.Add("TD002", PoErpNo);
                    dynamicParameters.Add("TD003", PoSeq);
                    var resultCOPTD = sqlConnection2.Query(sql, dynamicParameters);
                    if (resultCOPTD.Count() <= 0) throw new SystemException("採購單找不到,請重新確認");
                    foreach (var item1 in resultCOPTD)
                    {
                        if (item1.TD018 != "Y") throw new SystemException("【採購單:" + PoFull + "】非確認狀態,不可操作!!!");
                        if (item1.TD016 == "Y") throw new SystemException("【採購單:" + PoFull + "】已經結案!!!");
                        Double TD008 = Convert.ToDouble(item1.TD008);
                        Double TD015 = Convert.ToDouble(item1.TD015);
                        PoRate = Convert.ToDouble(item1.PoRate);
                        PoAvailableRate = Convert.ToDouble(item1.PoAvailableRate);
                        PoConversionQty = Convert.ToDouble(item1.PoConversionQty);
                        Double PoConversionSiQty = Convert.ToDouble(item1.PoConversionSiQty);
                        PoConversionAvailableQty = Convert.ToDouble(item1.PoConversionAvailableQty);
                        Double PoConversionAvailableSiQty = Convert.ToDouble(item1.PoConversionAvailableSiQty);

                        if (PoConversionQty == PoConversionSiQty + ConversionQty)
                        {
                            CloseStatus = "Y";
                        }
                        else if (PoConversionQty < PoConversionSiQty + ConversionQty)
                        {
                            throw new SystemException("【採購單:" + PoFull + "】已超過採購數量,目前已交數量:" + TD015 + ",本次進貨驗收數量:" + AcceptQty + ",不可操作!!!");
                        }
                    }
                    #endregion
                    #endregion

                    #region //確認段
                    #region //固定更新
                    #region //PURMA 廠商基本資料檔
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE PURMA SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MA023 = @MA023
                                    WHERE MA001 = @SupplierNo";
                    dynamicParameters.AddDynamicParams(
                    new
                    {
                        MODIFIER = UserNo,
                        MODI_DATE = dateNow,
                        MODI_TIME = timeNow,
                        MODI_AP = BaseHelper.ClientComputer(),
                        MODI_PRID = "BM",
                        MA023 = ReceiveDateErp,
                        SupplierNo
                    });
                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion

                    #region //PURMB 品號廠商單頭檔
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE PURMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MB009 = @MB009
                                    WHERE MB001 = @MB001
                                    AND MB002 = @MB002";
                    dynamicParameters.AddDynamicParams(
                    new
                    {
                        MODIFIER = UserNo,
                        MODI_DATE = dateNow,
                        MODI_TIME = timeNow,
                        MODI_AP = BaseHelper.ClientComputer(),
                        MODI_PRID = "BM",
                        MB009 = ReceiveDateErp,
                        MB001 = MtlItemNo,
                        MB002 = SupplierNo
                    });
                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion

                    #region //PURTD 採購單單身資料檔
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE PURTD SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TD015 = TD015 + @TD015,
                                    TD016 = @TD016,
                                    TD060 = TD060 + @TD060
                                    WHERE TD001 = @TD001
                                    AND TD002 = @TD002
                                    AND TD003 = @TD003";
                    dynamicParameters.AddDynamicParams(
                    new
                    {
                        MODIFIER = UserNo,
                        MODI_DATE = dateNow,
                        MODI_TIME = timeNow,
                        MODI_AP = BaseHelper.ClientComputer(),
                        MODI_PRID = "BM",
                        TD015 = Math.Round(ConversionQty / PoRate, 3, MidpointRounding.AwayFromZero),
                        TD016 = CloseStatus,
                        TD060 = Math.Round(ConversionAvailableQty / PoAvailableRate, 3, MidpointRounding.AwayFromZero),
                        TD001 = PoErpPrefix,
                        TD002 = PoErpNo,
                        TD003 = PoSeq
                    });
                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion

                    #region //PURTH 進貨單單身檔
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE PURTH SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TH030 = 'Y',
                                    TH038 = @UserNo,
                                    TH004 = @TH004,
                                    TH005 = @TH005,
                                    TH006 = @TH006,
                                    TH007 = @TH007,
                                    TH008 = @TH008,
                                    TH009 = @TH009,
                                    TH010 = @TH010,
                                    TH011 = @TH011,
                                    TH012 = @TH012,
                                    TH013 = @TH013,
                                    TH014 = @TH014,
                                    TH015 = @TH015,
                                    TH016 = @TH016,
                                    TH017 = @TH017,
                                    TH018 = @TH018,
                                    TH019 = @TH019,
                                    TH020 = @TH020,
                                    TH021 = @TH021,
                                    TH022 = @TH022,
                                    TH023 = @TH023,
                                    TH024 = @TH024,
                                    TH025 = @TH025,
                                    TH026 = @TH026,
                                    TH027 = @TH027,
                                    TH028 = @TH028,
                                    TH029 = @TH029,
                                    TH031 = @TH031,
                                    TH032 = @TH032,
                                    TH033 = @TH033,
                                    TH034 = @TH034,
                                    TH035 = @TH035,
                                    TH036 = @TH036,
                                    TH037 = @TH037,
                                    TH039 = @TH039,
                                    TH040 = @TH040,
                                    TH041 = @TH041,
                                    TH042 = @TH042,
                                    TH043 = @TH043,
                                    TH044 = @TH044,
                                    TH045 = @TH045,
                                    TH046 = @TH046,
                                    TH047 = @TH047,
                                    TH048 = @TH048,
                                    TH049 = @TH049,
                                    TH050 = @TH050,
                                    TH051 = @TH051,
                                    TH052 = @TH052,
                                    TH053 = @TH053,
                                    TH054 = @TH054,
                                    TH055 = @TH055,
                                    TH056 = @TH056,
                                    TH057 = @TH057,
                                    TH058 = @TH058,
                                    TH059 = @TH059,
                                    TH060 = @TH060,
                                    TH061 = @TH061,
                                    TH062 = @TH062,
                                    TH063 = @TH063,
                                    TH064 = @TH064,
                                    TH065 = @TH065,
                                    TH066 = @TH066,
                                    TH067 = @TH067,
                                    TH068 = @TH068,
                                    TH069 = @TH069,
                                    TH070 = @TH070,
                                    TH071 = @TH071,
                                    TH072 = @TH072,
                                    TH073 = @TH073,
                                    TH074 = @TH074,
                                    TH500 = @TH500,
                                    TH501 = @TH501,
                                    TH502 = @TH502,
                                    TH503 = @TH503,
                                    TH550 = @TH550,
                                    TH551 = @TH551,
                                    TH552 = @TH552,
                                    TH553 = @TH553,
                                    TH554 = @TH554,
                                    TH555 = @TH555,
                                    TH556 = @TH556,
                                    TH557 = @TH557,
                                    TH558 = @TH558,
                                    TH075 = @TH075,
                                    TH076 = @TH076,
                                    TH077 = @TH077,
                                    TH078 = @TH078,
                                    TH079 = @TH079,
                                    TH080 = @TH080,
                                    TH081 = @TH081,
                                    TH082 = @TH082,
                                    TH083 = @TH083,
                                    TH084 = @TH084,
                                    TH085 = @TH085,
                                    TH086 = @TH086,
                                    TH087 = @TH087,
                                    TH088 = @TH088,
                                    TH089 = @TH089,
                                    TH090 = @TH090,
                                    TH091 = @TH091,
                                    TH092 = @TH092            
                                    WHERE TH001 = @GrErpPrefix
                                    AND TH002 = @GrErpNo
                                    AND TH003 = @GrSeq";
                    dynamicParameters.AddDynamicParams(
                    new
                    {
                        MODIFIER = UserNo,
                        MODI_DATE = dateNow,
                        MODI_TIME = timeNow,
                        MODI_AP = BaseHelper.ClientComputer(),
                        MODI_PRID = "BM",
                        GrErpPrefix,
                        GrErpNo,
                        GrSeq = grseq,
                        UserNo,
                        TH004,
                        TH005,
                        TH006,
                        TH007,
                        TH008,
                        TH009,
                        TH010,
                        TH011,
                        TH012,
                        TH013,
                        TH014,
                        TH015,
                        TH016,
                        TH017,
                        TH018,
                        TH019,
                        TH020,
                        TH021,
                        TH022,
                        TH023,
                        TH024,
                        TH025,
                        TH026,
                        TH027,
                        TH028,
                        TH029,
                        TH031,
                        TH032,
                        TH033,
                        TH034,
                        TH035,
                        TH036,
                        TH037,
                        TH039,
                        TH040,
                        TH041,
                        TH042,
                        TH043,
                        TH044,
                        TH045,
                        TH046,
                        TH047,
                        TH048,
                        TH049,
                        TH050,
                        TH051,
                        TH052,
                        TH053,
                        TH054,
                        TH055,
                        TH056,
                        TH057,
                        TH058,
                        TH059,
                        TH060,
                        TH061,
                        TH062,
                        TH063,
                        TH064,
                        TH065,
                        TH066,
                        TH067,
                        TH068,
                        TH069,
                        TH070,
                        TH071,
                        TH072,
                        TH073,
                        TH074,
                        TH500,
                        TH501,
                        TH502,
                        TH503,
                        TH550,
                        TH551,
                        TH552,
                        TH553,
                        TH554,
                        TH555,
                        TH556,
                        TH557,
                        TH558,
                        TH075,
                        TH076,
                        TH077,
                        TH078,
                        TH079,
                        TH080,
                        TH081,
                        TH082,
                        TH083,
                        TH084,
                        TH085,
                        TH086,
                        TH087,
                        TH088,
                        TH089,
                        TH090,
                        TH091,
                        TH092
                    });
                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);

                    #endregion

                    #region //找是否有未確認單身，若無則更新單頭
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT (a.TH001 + '-' + a.TH002) GrErpFullNo, a.TH003 , a.TH030 detailConfirmStatus
                                    , b.TG013 headerConfirm
                                    FROM PURTH a
                                    INNER JOIN PURTG b ON (a.TH001 + '-' + a.TH002) = (b.TG001 + '-' + b.TG002)
                                    WHERE a.TH001 = @GrErpPrefix AND a.TH002 = @GrErpNo
                                    AND a.TH030 = 'N' 
                                    AND a.TH003 != @GrSeq";
                    dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                    dynamicParameters.Add("GrErpNo", GrErpNo);
                    dynamicParameters.Add("GrSeq", grseq);

                    var unConfirmerpresult = sqlConnection2.Query(sql, dynamicParameters);
                    #endregion

                    #region //更新 - 進貨單單頭
                    if (unConfirmerpresult.Count() <= 0)
                    {
                        #region //PURTG 進貨單單頭檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PURTG SET
                                        MODIFIER =@MODIFIER,
                                        MODI_DATE= @MODI_DATE,
                                        MODI_TIME= @MODI_TIME,
                                        MODI_AP= @MODI_AP,
                                        MODI_PRID= @MODI_PRID,
                                        FLAG= FLAG + 1,
                                        TG003 = @TG003,
                                        TG013 = 'Y'
                                        WHERE TG001 = @GrErpPrefix
                                        AND TG002 = @GrErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            GrErpPrefix,
                            GrErpNo,
                            TG003 = MaxReceiptDate
                        });
                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                        #endregion

                    }

                    #endregion

                    #region //INVMB 品號基本資料檔 庫存數量 庫存金額
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE INVMB SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MB064  = MB064 + @MB064,
                                    MB065  = MB065 + @MB065
                                    WHERE MB001 = @MB001";
                    dynamicParameters.AddDynamicParams(
                    new
                    {
                        MODIFIER = UserNo,
                        MODI_DATE = dateNow,
                        MODI_TIME = timeNow,
                        MODI_AP = BaseHelper.ClientComputer(),
                        MODI_PRID = "BM",
                        MB048 = CurrencyCode,
                        MB049 = Math.Round(TH018 / GrRate, 4, MidpointRounding.AwayFromZero),
                        MB050 = Math.Round(TH018 * Exchange / GrRate, 4, MidpointRounding.AwayFromZero),
                        MB064 = ConversionQty,
                        MB065 = TH047,
                        MB001 = MtlItemNo
                    });
                    rowsAffected = sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion
                    #endregion

                    #region //固定新增
                    #region //INVLA 異動明細資料檔
                    dynamicParameters = new DynamicParameters();
                    sql = @"INSERT INTO INVLA (COMPANY, CREATOR, USR_GROUP, 
                                    CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                    UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                    LA001, LA004, LA005, LA006, LA007,
                                    LA008, LA009, LA010, LA011, LA012,
                                    LA013, LA014, LA015, LA016, LA017,
                                    LA018, LA019, LA020, LA021, LA022,
                                    LA023, LA024, LA025, LA026, LA027,
                                    LA028, LA029, LA030, LA031, LA032,
                                    LA033, LA034)
                                    OUTPUT INSERTED.LA001
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                    @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                    '','','','','',0,0,0,0,0,
                                    @LA001, @LA004, @LA005, @LA006, @LA007,
                                    @LA008, @LA009, @LA010, @LA011, @LA012,
                                    @LA013, @LA014, @LA015, @LA016, @LA017,
                                    @LA018, @LA019, @LA020, @LA021, @LA022,
                                    0, 0, '', '', '',
                                    '', '','','','',
                                    '', '')";
                    dynamicParameters.AddDynamicParams(
                        new
                        {
                            COMPANY = ErpNo,
                            CREATOR = UserNo,
                            USR_GROUP,
                            CREATE_DATE = dateNow,
                            MODIFIER = "",
                            MODI_DATE = "",
                            FLAG = 1,
                            CREATE_TIME = timeNow,
                            CREATE_AP = BaseHelper.ClientComputer(),
                            CREATE_PRID = "BM",
                            LA001 = MtlItemNo,
                            LA004 = MaxReceiptDate,
                            LA005 = 1,
                            LA006 = GrErpPrefix,
                            LA007 = GrErpNo,
                            LA008 = grseq,
                            LA009 = InventoryNo,
                            LA010 = SupplierNo,
                            LA011 = ConversionQty,
                            LA012 = UnitCost,
                            LA013 = TH047,
                            LA014 = "1",
                            LA015 = "Y",
                            LA016 = LotNumber,
                            LA017 = TH047, //金額-材料 暫未確認 
                                    LA018 = 0,
                            LA019 = 0,
                            LA020 = 0,
                            LA021 = 0,
                            LA022 = Location != "" ? Location : "" //儲位尚未給
                                });
                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                    #endregion
                    #endregion

                    #region //第一次新增後續更新

                    #region //INVMC 品號庫別檔異動
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT MC012 FROM INVMC
                                    WHERE MC001 = @MC001
                                      AND MC002 = @MC002";
                    dynamicParameters.Add("MC001", MtlItemNo);
                    dynamicParameters.Add("MC002", InventoryNo);
                    var resultINVMC = sqlConnection2.Query(sql, dynamicParameters);

                    if (resultINVMC.Count() > 0)
                    {
                        #region //後續更新流程
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE INVMC SET
                                        MODIFIER =@MODIFIER,
                                        MODI_DATE= @MODI_DATE,
                                        MODI_TIME= @MODI_TIME,
                                        MODI_AP= @MODI_AP,
                                        MODI_PRID= @MODI_PRID,
                                        FLAG= FLAG + 1,
                                        MC007  = MC007 + @MC007,
                                        MC008  = MC008 + @MC008,
                                        MC012  = @MC012 
                                        WHERE MC001 = @MC001
                                        AND MC002 = @MC002";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            MC007 = ConversionQty,
                            MC008 = TH047,
                            MC012 = ReceiveDateErp,
                            MC001 = MtlItemNo,
                            MC002 = InventoryNo
                        });
                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                        #endregion
                    }
                    else
                    {
                        #region //第一次新增流程
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO INVMC (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        MC001, MC002, MC003, MC004, MC005, MC006, MC007, MC008, MC009, MC010, MC011,
                                        MC012, MC013, MC014, MC015)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        @MC001, @MC002, @MC003, @MC004, @MC005, @MC006, @MC007, @MC008, @MC009, @MC010, @MC011,
                                        @MC012, @MC013, @MC014, @MC015)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                COMPANY = ErpNo,
                                CREATOR = UserNo,
                                USR_GROUP,
                                CREATE_DATE = dateNow,
                                MODIFIER = "",
                                MODI_DATE = "",
                                FLAG = 1,
                                CREATE_TIME = timeNow,
                                CREATE_AP = BaseHelper.ClientComputer(),
                                CREATE_PRID = "BM",
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo,
                                MC003 = "",
                                MC004 = 0.000,
                                MC005 = 0.000,
                                MC006 = 0.000,
                                MC007 = ConversionQty,
                                MC008 = TH047,
                                MC009 = 0.000,
                                MC010 = 0.00000,
                                MC011 = "",
                                MC012 = ReceiveDateErp,
                                MC013 = "",
                                MC014 = 0.000,
                                MC015 = ""
                            });
                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                        #endregion

                    }
                    #endregion

                    #endregion

                    #region //品號有批號管理觸發
                    if (LotManagement != "N")
                    {
                        #region //固定新增
                        #region //INVLF 儲位批號異動明細資料檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO INVLF (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                        LF001, LF002, LF003, LF004, LF005,
                                        LF006, LF007, LF008, LF009, LF010,
                                        LF011, LF012, LF013)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        '','','','','',0,0,0,0,0,
                                        @LF001, @LF002, @LF003, @LF004, @LF005,
                                        @LF006, @LF007, @LF008, @LF009, @LF010,
                                        @LF011, @LF012, @LF013)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                COMPANY = ErpNo,
                                CREATOR = UserNo,
                                USR_GROUP,
                                CREATE_DATE = dateNow,
                                MODIFIER = "",
                                MODI_DATE = "",
                                FLAG = 1,
                                CREATE_TIME = timeNow,
                                CREATE_AP = BaseHelper.ClientComputer(),
                                CREATE_PRID = "BM",
                                LF001 = GrErpPrefix,
                                LF002 = GrErpNo,
                                LF003 = grseq,
                                LF004 = MtlItemNo,
                                LF005 = InventoryNo,
                                LF006 = Location.Length <= 0 ? "##########" : Location,
                                LF007 = LotNumber.Length <= 0 ? "####################" : LotNumber,
                                LF008 = 1,
                                LF009 = ReceiveDateErp,
                                LF010 = 1,
                                LF011 = ConversionQty,
                                LF012 = 0,
                                LF013 = SupplierNo
                            });

                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                        #endregion

                        #region //INVMF 品號批號資料單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                        MF001, MF002, MF003, MF004, MF005, 
                                        MF006, MF007, MF008, MF009, MF010, 
                                        MF013, MF014,
                                        MF016, MF017, MF502, MF503, MF504)
                                 VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        '','','','','',0,0,0,0,0,
                                        @MF001, @MF002, @MF003, @MF004, @MF005,
                                        @MF006, @MF007, @MF008, @MF009, @MF010,
                                        @MF013, @MF014,
                                        0, 0, 0, 0, 0)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                COMPANY = ErpNo,
                                CREATOR = UserNo,
                                USR_GROUP,
                                CREATE_DATE = dateNow,
                                MODIFIER = "",
                                MODI_DATE = "",
                                FLAG = 1,
                                CREATE_TIME = timeNow,
                                CREATE_AP = BaseHelper.ClientComputer(),
                                CREATE_PRID = "BM",
                                MF001 = MtlItemNo,
                                MF002 = LotNumber,
                                MF003 = ReceiveDateErp,
                                MF004 = GrErpPrefix,
                                MF005 = GrErpNo,
                                MF006 = grseq,
                                MF007 = InventoryNo,
                                MF008 = 1,
                                MF009 = "1",
                                MF010 = ConversionQty,
                                MF013 = "", //備註
                                        MF014 = 0 //包裝數量
                                    });

                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                        #endregion

                        #endregion

                        #region //只新增一次
                        #region //INVME 品號批號資料單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ME001,ME007 FROM INVME
                                        WHERE ME001 = @ME001
                                          AND ME002 = @ME002";
                        dynamicParameters.Add("ME001", MtlItemNo);
                        dynamicParameters.Add("ME002", LotNumber);
                        var resultINVME = sqlConnection2.Query(sql, dynamicParameters);
                        if (resultINVME.Count() <= 0)
                        {
                            #region //INVME 新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO INVME (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       ME001, ME002, ME003, ME004, ME005, 
                                       ME006, ME007, ME008, ME009, ME010,
                                       ME011, ME012, ME503, ME504, ME505 )
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @ME001, @ME002, @ME003, @ME004, @ME005, 
                                       @ME006, @ME007, @ME008, @ME009, @ME010,
                                       0, 0, 0, 0, 0 )";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    COMPANY = ErpNo,
                                    CREATOR = UserNo,
                                    USR_GROUP,
                                    CREATE_DATE = dateNow,
                                    MODIFIER = "",
                                    MODI_DATE = "",
                                    FLAG = 1,
                                    CREATE_TIME = timeNow,
                                    CREATE_AP = BaseHelper.ClientComputer(),
                                    CREATE_PRID = "BM",
                                    ME001 = MtlItemNo,
                                    ME002 = LotNumber,
                                    ME003 = ReceiveDateErp,
                                    ME004 = "",
                                    ME005 = GrErpPrefix,
                                    ME006 = GrErpNo,
                                    ME007 = "N",
                                    ME008 = "",
                                    ME009 = AvailableDate,
                                    ME010 = ReCheckDate
                                });
                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion

                        }
                        else
                        {
                            foreach (var item1 in resultINVME)
                            {
                                if (item1.ME007 != "N") throw new SystemException("品號:" + MtlItemNo + "批號:" + LotNumber + "已經結案了!!!");
                            }
                        }
                        #endregion

                        #endregion

                        #region //第一次新增後續更新
                        #region //INVMM 品號庫別儲位批號檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MM001 FROM INVMM
                                        WHERE MM001 = @MM001
                                          AND MM002 = @MM002
                                          AND MM003 = @MM003
                                          AND MM004 = @MM004";
                        dynamicParameters.Add("MM001", MtlItemNo);
                        dynamicParameters.Add("MM002", InventoryNo);
                        dynamicParameters.Add("MM003", Location != "" ? Location : "##########");
                        dynamicParameters.Add("MM004", LotNumber);
                        var resultINVMM = sqlConnection2.Query(sql, dynamicParameters);

                        if (resultINVMM.Count() > 0)
                        {
                            #region //後續更新流程
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMM SET
                                                  MODIFIER = @MODIFIER,
                                                  MODI_DATE = @MODI_DATE,
                                                  MODI_TIME = @MODI_TIME,
                                                  MODI_AP = @MODI_AP,
                                                  MODI_PRID = @MODI_PRID,
                                                  FLAG = FLAG + 1,
                                                  MM005 = MM005 + @MM005,
                                                  MM008 = @MM008
                                            WHERE MM001 = @MM001
                                              AND MM002 = @MM002
                                              AND MM003 = @MM003
                                              AND MM004 = @MM004";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MM005 = ConversionQty,
                                MM008 = ReceiveDateErp,
                                MM001 = MtlItemNo,
                                MM002 = InventoryNo,
                                MM003 = Location != "" ? Location : "##########",
                                MM004 = LotNumber,
                            });
                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //第一次新增流程
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO INVMM (COMPANY, CREATOR, USR_GROUP, 
                                            CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                            UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                            MM001, MM002, MM003, MM004, MM005, 
                                            MM006, MM007, MM008, MM009, MM010, 
                                            MM011, MM012, MM013, MM014)
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                            @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                            '','','','','',0,0,0,0,0,
                                            @MM001, @MM002, @MM003, @MM004, @MM005, 
                                            @MM006, @MM007, @MM008, @MM009, @MM010, 
                                            @MM011, @MM012, @MM013, @MM014)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    COMPANY = ErpNo,
                                    CREATOR = UserNo,
                                    USR_GROUP,
                                    CREATE_DATE = dateNow,
                                    MODIFIER = "",
                                    MODI_DATE = "",
                                    FLAG = 1,
                                    CREATE_TIME = timeNow,
                                    CREATE_AP = BaseHelper.ClientComputer(),
                                    CREATE_PRID = "BM",
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location.Length > 0 ? Location : "##########",
                                    MM004 = LotNumber.Length > 0 ? LotNumber : "####################",
                                    MM005 = ConversionQty,
                                    MM006 = 0, //庫存包裝數量,
                                            MM007 = "",
                                    MM008 = ReceiveDateErp,
                                    MM009 = "",
                                    MM010 = "", //備註
                                            MM011 = 0,
                                    MM012 = "",
                                    MM013 = 0,
                                    MM014 = 0
                                });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #endregion
                    }
                    #endregion

                    #region //進貨單來源為暫入單觸發
                    if (TrErpPrefix != "")
                    {
                        //#region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"UPDATE INVTF 
                        //        SET MODIFIER =@MODIFIER,
                        //       MODI_DATE= @MODI_DATE,
                        //       MODI_TIME= @MODI_TIME,
                        //       MODI_AP= @MODI_AP,
                        //       MODI_PRID= @MODI_PRID,
                        //       FLAG= FLAG + 1,
                        //        TF042  = @TF042
                        //        WHERE TF001 = @TF001
                        //        AND TF002 = @TF002";
                        //dynamicParameters.AddDynamicParams(
                        //new
                        //{
                        //    TF042 = ReceiveDateErp,
                        //    TF001 = TH032,
                        //    TF002 = TH033
                        //});
                        //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //#endregion

                        //#region //INVTG 暫出入轉撥單身檔 轉銷數量
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"UPDATE INVTG SET
                        //        TG020  = TG020 + @TG020,
                        //        TG046  = TG046 + @TG046,
                        //        TG054  = TG054 + @TG054,
                        //        TG024  = @TG024
                        //        WHERE TG001 = @TG001
                        //        AND TG002 = @TG002
                        //        AND TG003 = @TG003";
                        //dynamicParameters.AddDynamicParams(
                        //new
                        //{
                        //    TG020 = Quantity,
                        //    TG046 = FreeSpareQty,
                        //    TG054 = AvailableQty,
                        //    TG001 = TH032,
                        //    TG002 = TH033,
                        //    TG003 = TH034,
                        //    TG024 = ClosureStatusINVTG
                        //});
                        //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //#endregion
                    }
                    #endregion

                    #endregion
                    
                }
                #endregion
                #endregion

                #region //MES

                foreach (var item in GRDetailResult)
                {
                    #region  //若單頭未確認，再找所有單身 確認是否有未確認單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT (a.GrErpPrefix + '-' + a.GrErpNo) GrErpFullNo , a.GrDetailId, a.ConfirmStatus detailConfirmStatus, a.GrSeq
                                , b.GrId, b.ConfirmStatus headerConfirm
                                FROM SCM.GrDetail a
                                INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                WHERE a.GrId = @GrId 
                                and a.ConfirmStatus = 'N' 
                                and a.GrDetailId !=  @GrDetailId";
                    dynamicParameters.Add("GrId", item.GrId);
                    dynamicParameters.Add("GrDetailId", GrdetailId);

                    var NoConfirmGRDetailResult = sqlConnection.Query(sql, dynamicParameters);

                    if (NoConfirmGRDetailResult.Count() <= 0)
                    {
                        DateTime date = DateTime.ParseExact(MaxReceiptDate, "yyyyMMdd", null); 
                        var maxReceiptDate = date.ToString("yyyy-MM-dd"); // yyyy-MM-dd 格式

                        #region //Update SCM.GoodsReceipt
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GoodsReceipt SET
                                     ReceiptDate = @ReceiptDate,
                                    ConfirmStatus = @ConfirmStatus,
                                    TransferStatus = 'Y',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ReceiptDate = maxReceiptDate,
                                ConfirmStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                item.GrId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }
                    #endregion

                    #region //批號相關異動
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.GrSeq,a.LotNumber,a.MtlItemId,b.LotManagement,a.InventoryId,a.AcceptQty
                                FROM SCM.GrDetail a
                                INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                                WHERE a.GrDetailId = @GrDetailId";
                    dynamicParameters.Add("GrDetailId", GrdetailId);

                    var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item2 in resultRoMes)
                    {
                        //string GrSeq = item.GrSeq;
                        int MtlItemId = item2.MtlItemId;
                        string LotNumber = item2.LotNumber;
                        string LotManagement = item2.LotManagement;
                        int InventoryId = item2.InventoryId;
                        double AcceptQty = Convert.ToDouble(item2.AcceptQty);
                        if (LotManagement != "N")
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
                            foreach (var item1 in MtlItemResult)
                            {
                                if (item1.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //確認此品號是否已存在此批號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LotNumberId
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            dynamicParameters.Add("LotNumberNo", LotNumber);

                            var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                            if (LotNumberResult.Count() <= 0)
                            {
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
                                        LotNumberNo = LotNumber,
                                        Remark = "",
                                        CreateDate = LastModifiedDate,
                                        LastModifiedDate,
                                        CreateBy = UserId,
                                        LastModifiedBy = UserId
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion

                                int LotNumberId = -1;
                                foreach (var item1 in insertResult)
                                {
                                    LotNumberId = item1.LotNumberId;
                                }

                                #region //INSERT SCM.LnDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix,FromErpNo,FromSeq
                                    ,InventoryId,TransactionType,DocType,Quantity,Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix,@FromErpNo,@FromSeq
                                    ,@InventoryId,@TransactionType,@DocType,@Quantity,@Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LotNumberId,
                                        TransactionDate = ReceiptDate,
                                        FromErpPrefix = GrErpPrefix,
                                        FromErpNo = GrErpNo,
                                        FromSeq = grseq,
                                        InventoryId,
                                        TransactionType = 1,
                                        DocType = 1,
                                        Quantity = AcceptQty,
                                        Remark = "",
                                        CreateDate = LastModifiedDate,
                                        LastModifiedDate,
                                        CreateBy = UserId,
                                        LastModifiedBy = UserId
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion
                            }
                            else
                            {
                                int LotNumberId = -1;
                                foreach (var item1 in LotNumberResult)
                                {
                                    LotNumberId = item1.LotNumberId;
                                }

                                #region //INSERT SCM.LnDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix,FromErpNo,FromSeq
                                    ,InventoryId,TransactionType,DocType,Quantity,Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix,@FromErpNo,@FromSeq
                                    ,@InventoryId,@TransactionType,@DocType,@Quantity,@Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LotNumberId,
                                        TransactionDate = ReceiptDate,
                                        FromErpPrefix = GrErpPrefix,
                                        FromErpNo = GrErpNo,
                                        FromSeq = grseq,
                                        InventoryId,
                                        TransactionType = 1,
                                        DocType = 1,
                                        Quantity = AcceptQty,
                                        Remark = "",
                                        CreateDate = LastModifiedDate,
                                        LastModifiedDate,
                                        CreateBy = UserId,
                                        LastModifiedBy = UserId
                                    });
                                var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult1.Count();
                                #endregion
                            }
                            #endregion
                        }
                    }
                    #endregion


                    #region //Update SCM.GrDetail
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE SCM.GrDetail SET
                                    ConfirmStatus = @ConfirmStatus,
                                ConfirmUser = @ConfirmUser,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                    WHERE GrDetailId = @GrDetailId";
                    dynamicParameters.AddDynamicParams(
                        new
                        {
                            ConfirmStatus = "Y",
                            ConfirmUser = UserId,
                            LastModifiedDate,
                            LastModifiedBy = UserId,
                            GrdetailId
                        });

                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                    #endregion
                }

                #region /撈取進貨單單身資料=>更新採購單,暫入單
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.AcceptQty,a.AvailableQty,a.PoDetailId
                                FROM SCM.GrDetail a
                                WHERE a.GrDetailId = @GrDetailId";
                dynamicParameters.Add("GrDetailId", GrdetailId);

                var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item3 in resultGrDetail)
                {
                    #region //更新 - 採購單 已交數量,已交計價數量
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE SCM.PoDetail SET
                                    SiQty = SiQty + @SiQty,
                                    SiPriceQty = SiPriceQty + @SiPriceQty,
                                    ClosureStatus = @ClosureStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoErpPrefix = @PoErpPrefix AND PoErpNo = @PoErpNo AND PoSeq = @PoSeq";
                    dynamicParameters.AddDynamicParams(
                            new
                            {
                                SiQty = item3.AcceptQty,
                                SiPriceQty = item3.AvailableQty,
                                ClosureStatus = CloseStatus,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                PoErpPrefix,
                                PoErpNo,
                                PoSeq
                            });
                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                    #endregion

                    #region //更新 - 暫入單 轉進銷量,轉銷贈/備品量
                    //if (item.TsnDetailId > 0)
                    //{
                    //    dynamicParameters = new DynamicParameters();
                    //    sql = @"UPDATE SCM.TsnDetail SET
                    //            SaleQty = SaleQty + @SaleQty,
                    //            SaleFSQty = SaleFSQty + @SaleFSQty,
                    //            LastModifiedDate = @LastModifiedDate,
                    //            LastModifiedBy = @LastModifiedBy
                    //            WHERE TsnDetailId = @TsnDetailId";
                    //    dynamicParameters.AddDynamicParams(
                    //        new
                    //        {
                    //            SaleQty = item.Quantity,
                    //            SaleFSQty = item.FreeSpareQty,
                    //            LastModifiedDate,
                    //            LastModifiedBy,
                    //            TsnDetailId = item.TsnDetailId
                    //        });
                    //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                    //}
                    #endregion
                    #endregion
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

            }

            return jsonResponse.ToString();
        }
        #endregion
    }
}
