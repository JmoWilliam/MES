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
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace MESDA
{
    public class PrintLabelDA
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

        public PrintLabelDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];           
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
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

        #region//Get
        #region //GetLabelPrintMachine02 -- 取得標籤機資訊 -- Daiyi 2023.02.14
        public string GetLabelPrintMachine02(int LabelPrintId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.LabelPrintId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.LabelPrintNo, a.LabelPrintName, a.LabelPrintIp
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.LabelPrintMachine a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LabelPrintId", @" AND a.LabelPrintId = @LabelPrintId", LabelPrintId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LabelPrintId DESC";
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

        #region//WarehouseStorageForPcs 入庫標籤
        #region//GetBarcode --取得入庫產品條碼資訊
        public string GetBarcode(string BarcodeNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT d.MtlItemNo,d.MtlItemName,b.ModeId
                             FROM MES.Barcode a
                             INNER JOIN MES.ManufactureOrder b ON a.MoId=b.MoId
                             INNER JOIN MES.WipOrder c ON b.WoId=c.WoId
                             INNER JOIN PDM.MtlItem d ON c.MtlItemId=d.MtlItemId
                             WHERE a.BarcodeNo=@BarcodeNo";
                    dynamicParameters.Add("@BarcodeNo", BarcodeNo);
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

        #region//GetWorkStage --取得加工階段
        public string GetWorkStage(int ModeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT b.ProcessName,b.ProcessDesc
                             FROM MES.ProcessParameter a
                             INNER JOIN MES.Process b ON a.ProcessId=b.ProcessId
                             WHERE a.ModeId=@ModeId";
                    dynamicParameters.Add("@ModeId", ModeId);
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

        #region//OutsideIdentificationTag 中潤箱號標籤

        #region//GetLabelPrintMachine --取得取得標籤機
        public string GetLabelPrintMachine(string LabelPrintNo,int CompanyId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT LabelPrintNo,LabelPrintName
                            FROM MES.LabelPrintMachine"; 
                    if(CompanyId > 0)
                    {
                        sql += " WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                    }
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

        #region//GetErpCustData 取得產品客戶資訊
        public string GetErpCustData(string BarcodeNo)
        {
            try
            {
                string ErpNo = "";
                string ErpDbName = "";
                string MtlItemNo = "";
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT d.MtlItemNo,d.MtlItemName,b.ModeId
                             FROM MES.Barcode a
                             INNER JOIN MES.ManufactureOrder b ON a.MoId=b.MoId
                             INNER JOIN MES.WipOrder c ON b.WoId=c.WoId
                             INNER JOIN PDM.MtlItem d ON c.MtlItemId=d.MtlItemId
                             WHERE a.BarcodeNo=@BarcodeNo";
                    dynamicParameters.Add("@BarcodeNo", BarcodeNo);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    foreach (var item in result) {
                        MtlItemNo = item.MtlItemNo;
                    }


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
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MG003,MG005,MG001
                                FROM COPMG
                                WHERE MG002 LIKE @MtlItemNo";
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);
                        var erpResult = sqlConnection2.Query(sql, dynamicParameters);
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = erpResult
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

        #region //GetCompany -- 取得公司資料 -- Ben Ma 2022.04.11
        public string GetCompany(int CompanyId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, a.Telephone, a.Fax, a.Address
                        , a.AddressEn, ISNULL(a.LogoIcon, -1) LogoIcon, a.Status
                        , a.CompanyNo + ' ' + a.CompanyName CompanyWithNo";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CurrentCompany);                  

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

        #region//SunnyCustomerMoldLabel 舜宇出貨標籤

        #region//GetSunnyCustMoldLabel
        public string GetSunnyCustMoldLabel(string BarcodeNo,int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region//判斷BarcodeNo 是否為刻字條碼
                    if (BarcodeNo != "")
                    {
                        sql = @"SELECT f.MtlItemNo, f.MtlItemName, a.ItemValue AS 'BarcodeItem', c.BarcodeNo,a.CreateDate,g.MoItemPartId
                                FROM MES.BarcodeProcessAttribute a
                                    LEFT JOIN MES.BarcodeProcess b ON a.BarcodeProcessId=b.BarcodeProcessId
                                    LEFT JOIN MES.Barcode c ON c.BarcodeId=b.BarcodeId
                                    LEFT JOIN MES.ManufactureOrder d ON d.MoId=c.MoId
                                    LEFT JOIN MES.WipOrder e ON e.WoId=d.WoId
                                    LEFT JOIN PDM.MtlItem f ON f.MtlItemId=e.MtlItemId
                                    LEFT JOIN MES.MoItemPart g ON a.ItemValue=g.MoItemPartNo
                                WHERE a.ItemNo='Lettering' 
                                AND g.MoItemPartId = @BarcodeNo
                                ORDER BY a.CreateDate DESC";
                        dynamicParameters.Add("@BarcodeNo", BarcodeNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() == 1)
                        {
                            foreach (var item in result)
                            {
                                //將刻字條碼轉為產品條碼
                                BarcodeNo = item.BarcodeNo;
                            }
                        }
                        else
                        {
                            sql = @"SELECT f.MtlItemNo,f.MtlItemName,a.ItemValue AS 'BarcodeItem',c.BarcodeNo
                                    FROM MES.BarcodeProcessAttribute a
                                    INNER JOIN MES.BarcodeProcess b ON a.BarcodeProcessId=b.BarcodeProcessId
                                    INNER JOIN MES.Barcode c ON c.BarcodeId=b.BarcodeId
                                    INNER JOIN MES.ManufactureOrder d ON d.MoId=c.MoId
                                    INNER JOIN MES.WipOrder e ON e.WoId=d.WoId
                                    INNER JOIN PDM.MtlItem f ON f.MtlItemId=e.MtlItemId
                                    WHERE a.ItemNo='Lettering' AND c.BarcodeNo = @BarcodeNo 
                                    ORDER BY a.CreateDate DESC";
                            dynamicParameters.Add("@BarcodeNo", BarcodeNo);                          
                            result = sqlConnection.Query(sql, dynamicParameters);
                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result
                            });
                            #endregion
                        }
                    }
                    else {
                        sql = @"SELECT DISTINCT f.MtlItemNo, f.MtlItemName, a.ItemValue AS 'BarcodeItem', c.BarcodeNo
                                FROM MES.BarcodeProcessAttribute a
                                    INNER JOIN MES.BarcodeProcess b ON a.BarcodeProcessId=b.BarcodeProcessId
                                    INNER JOIN MES.Barcode c ON c.BarcodeId=b.BarcodeId
                                    INNER JOIN MES.ManufactureOrder d ON d.MoId=c.MoId
                                    INNER JOIN MES.WipOrder e ON e.WoId=d.WoId
                                    INNER JOIN PDM.MtlItem f ON f.MtlItemId=e.MtlItemId
                                WHERE a.ItemNo='Lettering' 
                                AND c.BarcodeNo IN (
                                    SELECT BarcodeNo
                                    FROM MES.Barcode
                                    WHERE d.MoId = @MoId
                                )
                                ORDER BY a.ItemValue DESC";
                        dynamicParameters.Add("@MoId", MoId);
                        var result = sqlConnection.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
                        });
                        #endregion
                    }
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

        #region SunnyCustMoldLabel 舜宇外箱標籤

        #region//GetSunnyBoxLabel
        public string GetSunnyBoxLabel(int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    sql = @"SELECT f.CustomerDwgNo,e.Edition,g.MtlItemNo,g.MtlItemName
                            FROM MES.ManufactureOrder a
                            LEFT JOIN MES.MoRouting b ON a.MoId=b.MoId
                            LEFT JOIN MES.RoutingItem c ON c.RoutingItemId=b.RoutingItemId
                            LEFT JOIN PDM.RdDesign d ON c.MtlItemId=d.MtlItemId
                            INNER JOIN PDM.CustomerCadControl e ON d.CustomerCadControlId=e.ControlId
                            INNER JOIN PDM.CustomerCad f ON f.CadId=e.CadId
                            INNER JOIN PDM.MtlItem g ON c.MtlItemId=g.MtlItemId
                             WHERE a.MoId = @MoId ";
                    dynamicParameters.Add("@MoId", MoId);
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

        #region//BlackLabel 黑物成型標籤        
        #region//GetBlackLabel
        public string GetBlackLabel(string BarcodeNo, int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    
                    if (BarcodeNo != "")
                    {
                        sql = @"SELECT a.MoId,c.MtlItemNo,c.MtlItemName,e.BarcodeNo
                                FROM MES.ManufactureOrder a
                                LEFT JOIN MES.WipOrder b ON a.WoId = b.WoId
                                LEFT JOIN PDM.MtlItem c ON c.MtlItemId = b.MtlItemId
                                LEFT JOIN MES.MoSetting d ON d.MoId = a.MoId
                                LEFT JOIN MES.BarcodePrint e ON e.MoSettingId = d.MoSettingId
                                WHERE 1=1 AND e.BarcodeNo=@BarcodeNo";
                        dynamicParameters.Add("@BarcodeNo", BarcodeNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
                        });
                        #endregion
                    }
                    else
                    {
                        sql = @"SELECT a.MoId,c.MtlItemNo,c.MtlItemName,e.BarcodeNo
                                FROM MES.ManufactureOrder a
                                LEFT JOIN MES.WipOrder b ON a.WoId = b.WoId
                                LEFT JOIN PDM.MtlItem c ON c.MtlItemId = b.MtlItemId
                                LEFT JOIN MES.MoSetting d ON d.MoId = a.MoId
                                LEFT JOIN MES.BarcodePrint e ON e.MoSettingId = d.MoSettingId
                                WHERE 1=1 AND a.MoId=@MoId";                        
                        dynamicParameters.Add("@MoId", MoId);
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
            }
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

        #region //GetPackageBarcodeLabel
        public string GetPackageBarcodeLabel(int PackageBarcodeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {


                    sql = @"SELECT  a.PackageBarcodeId LabBarcodeId, a.PackageBarcodeNo BarcodeNo,c.TotalCount
                            FROM MES.PackageBarcode a
                            INNER JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                            OUTER  APPLY(
	                            SELECT COUNT(x.PbrId) TotalCount from MES.PackageBarcodeReference x
	                            WHERE a.PackageBarcodeId = x.PackageBarcodeId
                            ) c
                            WHERE 1=1 AND a.PackageBarcodeId = @PackageBarcodeId";
                    dynamicParameters.Add("PackageBarcodeId", PackageBarcodeId);
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

        #region //GetLotBarcodeForLabel -- 標籤機取得批量條碼或包裝條碼 -- Shintokru 2023.05.23
        public string GetLotBarcodeForLabel(string MoId, string BarcodeNo, string BarcodeType, string PrintCnt)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    if (MoId.IndexOf('-') != -1)
                    {
                        if (MoId.IndexOf('(') == -1)
                        {
                            throw new SystemException("請連製令序號一起輸入! EX:5101-20220929001(1)");
                        }
                        else
                        {
                            string otherWoErpPrefix = MoId.Split('-')[0];
                            string tempOtherWoErpNo = MoId.Split('-')[1];
                            string otherWoErpNo = tempOtherWoErpNo.Split('(')[0];
                            string woSeq = tempOtherWoErpNo.Split('(')[1].Split(')')[0];

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.MoId
                                    FROM MES.WipOrder a
                                    INNER JOIN MES.ManufactureOrder b ON a.WoId = b.WoId
                                    WHERE a.WoErpPrefix = @WoErpPrefix AND a.WoErpNo = @WoErpNo
                                    AND b.WoSeq = @WoSeq";
                            dynamicParameters.Add("WoErpPrefix", otherWoErpPrefix);
                            dynamicParameters.Add("WoErpNo", otherWoErpNo);
                            dynamicParameters.Add("WoSeq", woSeq);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            if (result2.Count() <= 0) throw new SystemException("查無對應製令資料!");
                            foreach (var item in result2)
                            {
                                MoId = item.MoId.ToString();
                            }
                        }
                    }

                    #region //判斷製令生產模式
                    int ModeId = -1;
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.ModeId FROM  MES.ManufactureOrder a 
                            INNER JOIN MES.MoSetting b on a.MoId = b.MoId
                            INNER JOIN MES.BarcodePrint c on b.MoSettingId = c.MoSettingId
                            LEFT JOIN MES.Tray d on c.BarcodeNo = d.BarcodeNo
                            WHERE 1=1";
                    if (MoId.Length > 0)
                    {
                        sql += @" AND a.MoId = @MoId";
                        dynamicParameters.Add("MoId", MoId);
                    }
                    if (BarcodeNo.Length > 0)
                    {
                        sql += @" AND c.BarcodeNo = @BarcodeNo";
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);
                    }

                    var resultMode = sqlConnection.Query(sql, dynamicParameters);
                    if (resultMode.Count() > 0)
                    {
                        foreach (var item in resultMode)
                        {
                            ModeId = item.ModeId;
                        }
                    }
                    else{
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ModeId FROM  MES.ManufactureOrder a 
                            INNER JOIN MES.MoSetting b on a.MoId = b.MoId
                            INNER JOIN MES.BarcodePrint c on b.MoSettingId = c.MoSettingId
                            LEFT JOIN MES.Tray d on c.BarcodeNo = d.BarcodeNo
                            WHERE 1=1";
                        if (MoId.Length > 0)
                        {
                            sql += @" AND a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);
                        }
                        if (BarcodeNo.Length > 0)
                        {
                            sql += @" AND d.TrayNo = @TrayNo";
                            dynamicParameters.Add("TrayNo", BarcodeNo);
                        }
                        resultMode = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMode.Count() > 0)
                        {
                            foreach (var item in resultMode)
                            {
                                ModeId = item.ModeId;
                            }
                        }
                        else{
                            throw new SystemException("條碼查無對應生產模式資料!");
                        }
                    }
                    
                    #endregion

                    if (Convert.ToInt32(BarcodeType) == 0)
                    {
                        if (ModeId == 48)
                        {
                            sql = @"SELECT a.MoId,d.ItemNo,d.ItemValue,b.PrintId LabBarcodeId,b.BarcodeNo,b.BarcodeQty,f.PlanQty
                                    ,(f.WoErpPrefix + '-' + f.WoErpNo + '(' + CONVERT(VARCHAR(50), e.WoSeq)+')') WoErpFull,g.MtlItemName
                                    FROM MES.MoSetting a
                                    INNER JOIN MES.BarcodePrint b on a.MoSettingId = b.MoSettingId
                                    INNER JOIN MES.Barcode c on b.BarcodeNo = c.BarcodeNo
                                    INNER JOIN MES.BarcodeAttribute d on c.BarcodeId = d.BarcodeId
                                    INNER JOIN MES.ManufactureOrder e on a.MoId = e.MoId
                                    INNER JOIN MES.WipOrder f on e.WoId = f.WoId
                                    INNER JOIN PDM.MtlItem g on f.MtlItemId = g.MtlItemId
                                    WHERE 1=1
                                    AND d.ItemNo = 'Cavity'
                                    AND d.ItemValue is not null";
                            if (MoId.Length > 0)
                            {
                                sql += @" AND a.MoId = @MoId";
                            }
                            if (BarcodeNo.Length > 0)
                            {
                                sql += @" AND b.BarcodeNo = @BarcodeNo";
                            }
                            switch (PrintCnt)
                            {
                                case "Y":
                                    sql += @" AND b.PrintCnt > 0";
                                    break;
                                case "N":
                                    sql += @" AND b.PrintCnt = 0";
                                    break;
                                case "":
                                    break;
                            }
                        }
                        else
                        {
                            sql = @"SELECT a.MoId,b.PrintId LabBarcodeId,b.BarcodeNo,b.BarcodeQty,d.PlanQty
                                    ,(d.WoErpPrefix + '-' + d.WoErpNo + '(' + CONVERT(VARCHAR(50), c.WoSeq)+')') WoErpFull ,e.MtlItemName
                                    FROM MES.MoSetting a
                                    INNER JOIN MES.BarcodePrint b on a.MoSettingId = b.MoSettingId
                                    INNER JOIN MES.ManufactureOrder c on a.MoId = c.MoId
                                    INNER JOIN MES.WipOrder d on c.WoId = d.WoId
                                    INNER JOIN PDM.MtlItem e on d.MtlItemId = e.MtlItemId
                                    WHERE 1=1";
                            if (MoId.Length > 0)
                            {
                                sql += @" AND a.MoId = @MoId";
                            }
                            if (BarcodeNo.Length > 0)
                            {
                                sql += @" AND b.BarcodeNo = @BarcodeNo";
                            }
                            switch (PrintCnt)
                            {
                                case "Y":
                                    sql += @" AND b.PrintCnt > 0";
                                    break;
                                case "N":
                                    sql += @" AND b.PrintCnt = 0";
                                    break;
                                case "":
                                    break;
                            }
                        }
                    }
                    else if (Convert.ToInt32(BarcodeType) == 1)
                    {
                        if (BarcodeNo == "") throw new SystemException("如果要撈取Tray條碼,條碼不可以為空!");
                        sql = @"SELECT b.MoId,,a.TrayNo BarcodeNo ,d.PlanQty
                                ,(d.WoErpPrefix + '-' + d.WoErpNo + '(' + CONVERT(VARCHAR(50), c.WoSeq)+')') WoErpFull ,e.MtlItemName
                                FROM MES.Tray a
                                LEFT JOIN MES.Barcode b on a.BarcodeNo = b.BarcodeNo
                                INNER JOIN MES.ManufactureOrder c on b.MoId = c.MoId
                                INNER JOIN MES.WipOrder d on c.WoId = d.WoId
                                INNER JOIN PDM.MtlItem e on d.MtlItemId = e.MtlItemId
                                WHERE 1=1
                                AND a.TrayNo = @BarcodeNo";
                    }
                    else if (Convert.ToInt32(BarcodeType) == 2)
                    {
                        sql = @"SELECT  b.MoId,a.PackageBarcodeId LabBarcodeId, a.PackageBarcodeNo BarcodeNo,c.total
                                ,(d.WoErpPrefix + '-' + d.WoErpNo + '(' + CONVERT(VARCHAR(50), c.WoSeq)+')') WoErpFull ,e.MtlItemName ,d.PlanQty 
                                FROM MES.PackageBarcode a
                                INNER JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                                INNER JOIN MES.ManufactureOrder c on a.MoId = c.MoId
                                INNER JOIN MES.WipOrder d on c.WoId = d.WoId
                                INNER JOIN PDM.MtlItem e on d.MtlItemId = e.MtlItemId
                                OUTER  APPLY(
	                                SELECT COUNT(x.PbrId) total from MES.PackageBarcodeReference x
	                                WHERE a.PackageBarcodeId = x.PackageBarcodeId
                                ) c
                                WHERE 1=1";
                        if (MoId.Length > 0)
                        {
                            sql += @" AND b.MoId = @MoId";
                        }
                        if (BarcodeNo.Length > 0)
                        {
                            sql += @" AND a.PackageBarcodeNo = @BarcodeNo";
                        }
                    }

                    dynamicParameters.Add("@MoId", MoId);
                    dynamicParameters.Add("@BarcodeNo", BarcodeNo);
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

        #region //GetMtlItem --取得品號
        public string GetMtlItem(string MtlItemNo, string MtlItemNoList)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷品號資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec
                            FROM PDM.MtlItem
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (MtlItemNo.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND MtlItemNo = @MtlItemNo", MtlItemNo);
                    }
                    if (MtlItemNoList.Length > 0) {
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNoList", @" AND MtlItemNo IN @MtlItemNoList", MtlItemNoList.Split(','));
                    }

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("品號資料錯誤!");
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

        #region //UpdateLotPrintCnt -- 更新批量條碼列印次數 -- Shintokru 2023.06.01
        public string UpdateLotPrintCnt(string PrintType, int MoId, string BarcodeNo, string BarcodeNoStr)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string PrintIdStr = "";

                        switch (PrintType)
                        {
                            case "PM":
                                #region //製令模式
                                #region //找出製令條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.PrintId
                                        FROM MES.MoSetting a
                                        INNER JOIN MES.BarcodePrint b on a.MoSettingId = b.MoSettingId
                                        WHERE a.MoId = @MoId";
                                dynamicParameters.Add("MoId", MoId);
                                var result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() <= 0) throw new SystemException("資料庫找不到該製令，請重新確認!");
                                foreach (var item in result1)
                                {
                                    if (PrintIdStr != "")
                                    {
                                        PrintIdStr += "," + item.PrintId;
                                    }
                                    else
                                    {
                                        PrintIdStr += "in (" + item.PrintId;
                                    }
                                }
                                #endregion
                                #endregion
                                break;
                            case "PB":
                                #region //判斷批量條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 PrintId
                                        FROM MES.BarcodePrint
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() <= 0) throw new SystemException("資料庫找不到該批量條碼資料，請重新確認!");
                                foreach (var item2 in result2)
                                {
                                   PrintIdStr += "in (" + item2.PrintId;
                                }
                                #endregion
                                break;
                            case "PT":
                                string[] BarcodeArr = BarcodeNoStr.Split(',');
                                foreach (var item in BarcodeArr)
                                {
                                    #region //判斷暫存的批量條碼資料是否正確
                                    int PrintId = -1;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 PrintId
                                            FROM MES.BarcodePrint
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.Add("BarcodeNo", item);

                                    var result3 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result3.Count() <= 0) throw new SystemException("資料庫找不到該批量條碼資料，請重新確認!");
                                    foreach (var item2 in result3)
                                    {
                                        PrintId = item2.PrintId;
                                        if (PrintIdStr != "")
                                        {
                                            PrintIdStr += "," + item2.PrintId;
                                        }
                                        else
                                        {
                                            PrintIdStr += "in (" + item2.PrintId;

                                        }
                                    }
                                    #endregion
                                }
                                break;
                            default:
                                throw new SystemException("找不到列印的模式,請重新確認");
                                break;
                        }
                        PrintIdStr = PrintIdStr + ")";

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.BarcodePrint SET
                                PrintCnt = PrintCnt + 1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PrintId " + PrintIdStr;
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                PrintIdStr
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

                        transactionScope.Complete();
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

        #region //量測標籤
        public string GetQCBarcodeLabel(int QcRecordId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //找出製令條碼資料是否正確
                    int MtlItemId = -1;
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT ISNULL(a.MtlItemId,-1) MtlItemId
                            FROM MES.QcRecord a                            
                            WHERE a.QcRecordId = @QcRecordId";
                    dynamicParameters.Add("QcRecordId", QcRecordId);
                    var result1 = sqlConnection.Query(sql, dynamicParameters);
                    if (result1.Count() <= 0) throw new SystemException("查無送測單據資料，請重新確認!");
                    foreach (var item in result1)
                    {
                        MtlItemId = item.MtlItemId;
                    }
                    #endregion

                   
                    if (MtlItemId==-1) {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId RID, a.Remark, c.WoErpPrefix + '-' + c.WoErpNo MOID, d.MtlItemName, d.MtlItemNo,a.Remark
                            , e.UserNo + '-' + e.UserName QcUser, f.DepartmentName QcDepartment,a.LastModifiedDate QcDate
                            FROM MES.QcRecord a
                            LEFT JOIN MES.ManufactureOrder b ON b.MoId = a.MoId
                            LEFT JOIN MES.WipOrder c ON c.WoId = b.WoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = c.MtlItemId
							LEFT JOIN BAS.[User] e ON a.CreateBy = e.UserId
							LEFT JOIN BAS.Department f ON e.DepartmentId = f. DepartmentId
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
                    else {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId RID, a.Remark, c.WoErpPrefix + '-' + c.WoErpNo MOID, d.MtlItemName, d.MtlItemNo,a.Remark
                            , e.UserNo + '-' + e.UserName QcUser, f.DepartmentName QcDepartment,a.LastModifiedDate QcDate
                            FROM MES.QcRecord a
                            LEFT JOIN MES.ManufactureOrder b ON b.MoId = a.MoId
                            LEFT JOIN MES.WipOrder c ON c.WoId = b.WoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
							LEFT JOIN BAS.[User] e ON a.CreateBy = e.UserId
							LEFT JOIN BAS.Department f ON e.DepartmentId = f. DepartmentId
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
            }
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

        public string GetLotQCBarcodeLabel(string QcRecordIds)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    string[] QcRecordIdsArray = QcRecordIds.Split(',');
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    foreach (var qcRecordId in QcRecordIdsArray)
                    {
                        #region //找出製令條碼資料是否正確
                        int MtlItemId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(a.MtlItemId,-1) MtlItemId
                            FROM MES.QcRecord a                            
                            WHERE a.QcRecordId = @QcRecordId";
                        dynamicParameters.Add("QcRecordId", qcRecordId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() <= 0) throw new SystemException("查無送測單據資料，請重新確認!");
                        foreach (var item in result1)
                        {
                            MtlItemId = item.MtlItemId;
                        }
                        #endregion


                        if (MtlItemId == -1)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId RID, a.Remark, c.WoErpPrefix + '-' + c.WoErpNo MOID, d.MtlItemName, d.MtlItemNo,a.Remark
                            , e.UserNo + '-' + e.UserName QcUser, f.DepartmentName QcDepartment
                            FROM MES.QcRecord a
                            LEFT JOIN MES.ManufactureOrder b ON b.MoId = a.MoId
                            LEFT JOIN MES.WipOrder c ON c.WoId = b.WoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = c.MtlItemId
							LEFT JOIN BAS.[User] e ON a.CreateBy = e.UserId
							LEFT JOIN BAS.Department f ON e.DepartmentId = f. DepartmentId
                            WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", qcRecordId);
                            result = sqlConnection.Query(sql, dynamicParameters);

                            
                        }
                        else
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId RID, a.Remark, c.WoErpPrefix + '-' + c.WoErpNo MOID, d.MtlItemName, d.MtlItemNo,a.Remark
                            , e.UserNo + '-' + e.UserName QcUser, f.DepartmentName QcDepartment
                            FROM MES.QcRecord a
                            LEFT JOIN MES.ManufactureOrder b ON b.MoId = a.MoId
                            LEFT JOIN MES.WipOrder c ON c.WoId = b.WoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
							LEFT JOIN BAS.[User] e ON a.CreateBy = e.UserId
							LEFT JOIN BAS.Department f ON e.DepartmentId = f. DepartmentId
                            WHERE a.QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", qcRecordId);
                            result = sqlConnection.Query(sql, dynamicParameters);

                            
                        }
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

        #region //UpdateBarcodePrintCount -- 更新條碼列印次數 -- Ann 2023-08-21
        public string UpdateBarcodePrintCount(string BarcodeNoList)
        {
            try
            {
                if (BarcodeNoList.Length <= 0) throw new SystemException("【條碼】列表不能為空!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string[] BarcodeList = BarcodeNoList.Split(',');

                        int rowsAffected = 0;
                        foreach (var barcodeNo in BarcodeList)
                        {
                            #region //判斷條碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Barcode a 
                                    WHERE a.BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", barcodeNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("條碼資料錯誤!!");
                            #endregion

                            #region //UPDATE MES.Barcode
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.Barcode SET
                                    PrintCount = PrintCount + 1,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    BarcodeNo = barcodeNo
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //若BarcodePrint有，則同步UPDATE MES.BarcodePrint
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.BarcodePrint SET
                                    PrintCnt = PrintCnt + 1,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    BarcodeNo = barcodeNo
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

        #region//取得條碼的公司別
        public string GetBarcodeCompany(string MoId, string BarcodeNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();                   
                    if (MoId.Length>0) {
                        if (MoId.IndexOf('-') != -1)
                        {
                            if (MoId.IndexOf('(') == -1)
                            {
                                throw new SystemException("請連製令序號一起輸入! EX:5101-20220929001(1)");
                            }
                            else
                            {
                                string otherWoErpPrefix = MoId.Split('-')[0];
                                string tempOtherWoErpNo = MoId.Split('-')[1];
                                string otherWoErpNo = tempOtherWoErpNo.Split('(')[0];
                                string woSeq = tempOtherWoErpNo.Split('(')[1].Split(')')[0];

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 c.CompanyNo
                                     FROM MES.ManufactureOrder a
                                     INNER JOIN MES.WipOrder b ON a.WoId=b.WoId
                                     INNER JOIN BAS.Company c ON b.CompanyId=c.CompanyId
                                    WHERE a.WoErpPrefix = @WoErpPrefix AND a.WoErpNo = @WoErpNo
                                    AND b.WoSeq = @WoSeq";
                                dynamicParameters.Add("WoErpPrefix", otherWoErpPrefix);
                                dynamicParameters.Add("WoErpNo", otherWoErpNo);
                                dynamicParameters.Add("WoSeq", woSeq);                              
                            }
                        }
                        else {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT  TOP 1 c.CompanyNo
                                     FROM MES.ManufactureOrder a
                                     INNER JOIN MES.WipOrder b ON a.WoId=b.WoId
                                     INNER JOIN BAS.Company c ON b.CompanyId=c.CompanyId
                                    WHERE a.MoId = @MoId ";
                            dynamicParameters.Add("MoId", MoId);                          
                        }
                    } else if (BarcodeNo.Length>0) {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT c.CompanyNo
                             FROM MES.ManufactureOrder a
                             INNER JOIN MES.WipOrder b ON a.WoId=b.WoId
                             INNER JOIN BAS.Company c ON b.CompanyId=c.CompanyId
                             INNER JOIN MES.Barcode d ON d.MoId=a.MoId
                             WHERE d.BarcodeNo=@BarcodeNo";
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);                       
                    }
                    else {
                        throw new SystemException("查無製令資料!");
                    }

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("查無此條碼對應製令資料!");

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

        #region//GetEtergeLotBarcode --取得紘立批量條碼資訊
        public string GetEtergeLotBarcode(string MoId, string BarcodeNo, string BarcodeType, string PrintCnt)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    //#region //Response
                    //jsonResponse = JObject.FromObject(new
                    //{
                    //    status = "success",
                    //    data = result
                    //});
                    //#endregion
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

        #region //CheckBarcodePrint -- 確認批量條碼資料 -- Ann 2024-04-11
        public string CheckBarcodePrint(string BarcodeNoList)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string[] BarcodeNoArray = BarcodeNoList.Split(',');
                    foreach (var barcodeNo in BarcodeNoArray)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrintId, a.BarcodeNo, a.BarcodeQty, a.MoSettingId, a.PrintCnt, a.PrintStatus
                                FROM MES.BarcodePrint a 
                                INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                WHERE a.BarcodeNo = @BarcodeNo";
                        dynamicParameters.Add("BarcodeNo", barcodeNo);

                        sql += @" ORDER BY b.BarcodeId";

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("批量條碼【" + barcodeNo + "】資料錯誤!!");
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = ""
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

        #region //GetBarcodePrint -- 取得批量條碼資料 -- Ann 2024-04-12
        public string GetBarcodePrint(int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PrintId, a.BarcodeNo, a.BarcodeQty, a.MoSettingId, a.PrintCnt, a.PrintStatus
                            FROM MES.BarcodePrint a 
                            INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoId", @" AND b.MoId = @MoId", MoId);

                    sql += @" ORDER BY b.BarcodeId";

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

        #region //GetLotLabel 批號條碼標籤列印
        public string GetLotLabel(int LotNumberId, string LnBarcodeNoList)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    string[] BarcodeNoArray = LnBarcodeNoList.Split(',');
                    foreach (var barcodeNo in BarcodeNoArray)
                    {
                        sql = @"SELECT  a.LnBarcodeId, b.LotNumberId, b.LotNumberNo, a.BarcodeNo, c.BarcodeQty, c.BarcodeStatus,  FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, d.UserNo, d.UserName
                            FROM SCM.LnBarcode a
                         INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                         INNER JOIN MES.Barcode c ON a.BarcodeNo = c.BarcodeNo
                         INNER JOIN BAS.[User] d ON a.CreateBy = d.UserId
                            WHERE 1=1 AND a.LotNumberId = @LotNumberId AND a.BarcodeNo = @BarcodeNo";
                        dynamicParameters.Add("LotNumberId", LotNumberId);
                        dynamicParameters.Add("BarcodeNo", barcodeNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("條碼【" + barcodeNo + "】資料錯誤!!");

                    }


                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = ""
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

        #region//GetBarcodeAttribute --取得條碼屬性
        public string GetBarcodeAttribute(string BarcodeNo, string ItemNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.ItemValue
                            FROM MES.BarcodeAttribute a 
                            INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                            WHERE b.BarcodeNo = @BarcodeNo
                            AND a.ItemNo = @ItemNo";
                    dynamicParameters.Add("ItemNo", ItemNo);
                    dynamicParameters.Add("BarcodeNo", BarcodeNo);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0)
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "null",
                            data = ""
                        });
                        #endregion
                    }
                    else
                    {
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

        #region //GetBarcodeandOrder -- 取得條碼及製令資料 -- WuTc 2024-12-10
        public string GetBarcodeandOrder(string BarcodeNo, int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.MoId
                            , Bp.BarcodeNo
                            , f.MtlItemId, f.MtlItemNo, f.MtlItemName
                            , e.PlanQty, g.PcPromiseDate
                            , (e.WoErpPrefix + '-' + e.WoErpNo + '(' + CONVERT(varchar, a.WoSeq) + ')') AS WoErpFull
                            FROM MES.ManufactureOrder a 
                            OUTER APPLY (
	                            SELECT bp.PrintId, bp.BarcodeNo, bp.BarcodeQty, bp.MoSettingId, bp.PrintCnt, bp.PrintStatus 
	                            FROM MES.MoSetting ms
	                            INNER JOIN MES.BarcodePrint bp ON ms.MoSettingId = bp.MoSettingId
	                            WHERE a.MoId = ms.MoId
                            ) Bp
                            INNER JOIN MES.WipOrder e ON a.WoId = e.WoId
                            INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
                            LEFT JOIN SCM.SoDetail g ON e.SoDetailId = g.SoDetailId
                            WHERE Bp.BarcodeNo != ''";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BarcodeNo", @" AND Bp.BarcodeNo = @BarcodeNo", BarcodeNo);

                    sql += @" UNION ALL 
                            SELECT DISTINCT a.MoId
                            , Bp.BarcodeNo
                            , f.MtlItemId, f.MtlItemNo, f.MtlItemName
                            , e.PlanQty, g.PcPromiseDate
                            , (e.WoErpPrefix + '-' + e.WoErpNo + '(' + CONVERT(varchar, a.WoSeq) + ')') AS WoErpFull
                            FROM MES.ManufactureOrder a 
                            OUTER APPLY (
	                            SELECT mbr.BarcodeNo FROM MES.MrWipOrder mo
	                            INNER JOIN MES.MrDetail md ON mo.MrId = md.MrId AND md.MoId = a.MoId
	                            LEFT JOIN MES.MrBarcodeRegister mbr ON md.MrDetailId = mbr.MrDetailId
	                            WHERE mo.MoId = a.MoId
                            ) Bp
                            INNER JOIN MES.WipOrder e ON a.WoId = e.WoId
                            INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
                            LEFT JOIN SCM.SoDetail g ON e.SoDetailId = g.SoDetailId
                            WHERE Bp.BarcodeNo != ''";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BarcodeNo", @" AND Bp.BarcodeNo = @BarcodeNo", BarcodeNo);

                    sql += @" ORDER BY a.MoId, Bp.BarcodeNo";

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
        #region //AddBarcodePrint -- 新增批量條碼(返修條碼) -- Ann 2023-11-21
        public string AddBarcodePrint(string WoErpFullNo, int BarcodeQty, int BarcodeCount)
        {
            try
            {
                if (BarcodeQty <= 0) throw new SystemException("條碼數量不能等於或小於0!!");
                if (BarcodeCount <= 0) throw new SystemException("條碼產生次數不能等於或小於0!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認製令資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoId
                                , c.MoSettingId, c.BarcodePrefix, c.SequenceLen, c.BarcodePostfix
                                , d.MoProcessId
                                FROM MES.ManufactureOrder a 
                                INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                INNER JOIN MES.MoSetting c ON a.MoId = c.MoId
                                INNER JOIN MES.MoProcess d ON a.MoId = d.MoId AND d.SortNumber = 1
                                WHERE b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(nvarchar(10), a.WoSeq) + ')' = @WoErpFullNo";
                        dynamicParameters.Add("WoErpFullNo", WoErpFullNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製令資料有誤!");

                        int MoSettingId = -1;
                        string BarcodePrefix = "";
                        int SequenceLen = -1;
                        string BarcodePostfix = "";
                        int MoProcessId = -1;
                        int MoId = -1;
                        foreach (var item in result)
                        {
                            MoSettingId = item.MoSettingId;
                            BarcodePrefix = "R-" + item.BarcodePrefix;
                            SequenceLen = item.SequenceLen;
                            BarcodePostfix = item.BarcodePostfix;
                            MoProcessId = item.MoProcessId;
                            MoId = item.MoId;
                        }
                        #endregion

                        string seq = "";
                        seq = GetNewBarcodeNo(BarcodePrefix, SequenceLen, BarcodePostfix).PadLeft(SequenceLen, '0');
                        string BarcodeNo = BarcodePrefix + seq + BarcodePostfix;

                        int rowsAffected = 0;
                        for (int i = 1; i <= BarcodeCount; i++)
                        {
                            if (i > 1)
                            {
                                seq = (Convert.ToInt32(seq) + 1).ToString();
                                seq = seq.PadLeft(SequenceLen, '0');
                                BarcodeNo = BarcodePrefix + seq + BarcodePostfix;
                            }

                            #region //INSERT MES.BarcodePrint
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.BarcodePrint (BarcodeNo, BarcodeQty, MoSettingId, ParentBarcode, PrintCnt, PrintStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@BarcodeNo, @BarcodeQty, @MoSettingId, @ParentBarcode, @PrintCnt, @PrintStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BarcodeNo,
                                    BarcodeQty,
                                    MoSettingId,
                                    ParentBarcode = "",
                                    PrintCnt = 0,
                                    PrintStatus = "P",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

                            #region //INSERT MES.Barcode
                            sql = @"INSERT INTO MES.Barcode (BarcodeNo, CurrentMoProcessId, NextMoProcessId, MoId, BarcodeQty, CurrentProdStatus, BarcodeStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.BarcodeId
                                    VALUES (@BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @MoId, @BarcodeQty, @CurrentProdStatus, @BarcodeStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BarcodeNo,
                                    CurrentMoProcessId = MoProcessId,
                                    NextMoProcessId = MoProcessId,
                                    MoId,
                                    BarcodeQty,
                                    CurrentProdStatus = "P",
                                    BarcodeStatus = "1",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
                        }

                        #region //取得最大序列號條碼
                        string GetNewBarcodeNo(string thisBarcodePrefix, int thisSequenceLen, string thisBarcodePostfix)
                        {
                            string NewBarcodeNo = "";
                            #region //找此BARCODE格式最大號
                            #region //組模糊查詢字串
                            string SearchLikeStrig = "";
                            for (int j = 1; j <= thisSequenceLen; j++)
                            {
                                SearchLikeStrig += "_";
                            }
                            #endregion

                            string queryModal = "";
                            if (thisBarcodePrefix.IndexOf("[") != -1)
                            {
                                char[] prefixList = thisBarcodePrefix.ToCharArray();

                                for (int i = 0; i < prefixList.Length; i++)
                                {
                                    if (prefixList[i] == '[') queryModal += "|";
                                    queryModal += prefixList[i];
                                }

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MAX(BarcodeNo) MaxBarcodeNo
                                            FROM MES.BarcodePrint
                                            WHERE BarcodeNo LIKE '%" + queryModal + SearchLikeStrig + thisBarcodePostfix + "%' ESCAPE '|'";
                            }
                            else if (thisBarcodePostfix.IndexOf("[") != -1)
                            {
                                char[] postfixList = thisBarcodePostfix.ToCharArray();

                                for (int i = 0; i < postfixList.Length; i++)
                                {
                                    if (postfixList[i] == '[') queryModal += "|";
                                    queryModal += postfixList[i];
                                }

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MAX(BarcodeNo) MaxBarcodeNo
                                            FROM MES.BarcodePrint
                                            WHERE BarcodeNo LIKE '%" + thisBarcodePrefix + SearchLikeStrig + queryModal + "%' ESCAPE '|'";
                            }
                            else
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MAX(BarcodeNo) MaxBarcodeNo
                                            FROM MES.BarcodePrint
                                            WHERE BarcodeNo LIKE '%" + thisBarcodePrefix + SearchLikeStrig + thisBarcodePostfix + "%'";
                            }

                            var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in BarcodeResult)
                            {
                                string MaxBarcodeNo = item.MaxBarcodeNo;
                                if (item.MaxBarcodeNo == null)
                                {
                                    seq = "1".PadLeft(SequenceLen, '0');
                                }
                                else
                                {
                                    #region //拆解BarcodeNo順序
                                    string thisSeq = "";
                                    if (thisBarcodePostfix != "")
                                    {
                                        thisSeq = MaxBarcodeNo.Replace(thisBarcodePrefix, "").Replace(thisBarcodePostfix, "");
                                    }
                                    else
                                    {
                                        thisSeq = MaxBarcodeNo.Replace(thisBarcodePrefix, "");
                                    }
                                    seq = (Convert.ToInt32(thisSeq) + 1).ToString().PadLeft(SequenceLen, '0');
                                    #endregion
                                }

                                NewBarcodeNo = thisBarcodePrefix + seq + thisBarcodePostfix;
                            }
                            #endregion

                            return seq;
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
}
