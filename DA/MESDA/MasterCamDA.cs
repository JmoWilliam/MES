using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Transactions;
using System.Web;

namespace MESDA
{
    public class MasterCamDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public string CreateUserNo = "";
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public SqlQuery sqlQuery = new SqlQuery();

        public MasterCamDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];            
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;
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

        #region//Get

        #region //GetUnitOfMeasure -- 取得單位資料 -- Ding 2022.10.09
        public string GetUnitOfMeasure(int UomId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT UomId,CompanyId ,UomNo,UomName
                            FROM PDM.UnitOfMeasure
                            WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND CompanyId = @CompanyId", CurrentCompany);
                    if (UomId > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND UomId= @UomId", UomId);
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

        #region //GetMoRdDesign -- 取得MES製令設計圖資料 -- Ding 2022-10-13
        public string GetMoRdDesign(int MoId, string OtherInfo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    #region //先檢查來源是否為BarcodeNo
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MoId
                            FROM MES.Barcode a
                            WHERE a.BarcodeNo = @BarcodeNo";
                    dynamicParameters.Add("BarcodeNo", OtherInfo);

                    var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item in BarcodeResult)
                    {
                        MoId = item.MoId;
                    }
                    #endregion

                    #region //檢查來源是否為WoErpNo
                    if (MoId==-1) {
                        if (OtherInfo.IndexOf("-") != -1)
                        {
                            if (OtherInfo.IndexOf("(") == -1) throw new SystemException("請連製令序號一起輸入! EX:5101-20220929001(1)");
                            else
                            {
                                string otherWoErpPrefix = OtherInfo.Split('-')[0];
                                string tempOtherWoErpNo = OtherInfo.Split('-')[1];
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
                                    MoId = item.MoId;
                                }
                            }
                        }
                    }                   
                    #endregion

                    sqlQuery.mainKey = "a.MoRoutingId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MoId, a.RoutingItemId, a.SortNumber,g.WoErpNo,f.WoSeq
                        , b.ControlId, b.MtlItemId, b.RoutingItemConfirm
                        , c.RoutingId, c.RoutingConfirm, c.ModeId, c.RoutingName
                        , d.ModeNo, d.ModeName, e.DesignDate, e.Edition,e.[Version]
                        ,(
                            SELECT  xa.FileId, (xa.[FileName] + xa.FileExtension) FileInfo
                            FROM BAS.[File] xa
                            WHERE e.Cad2DFile=xa.FileId
                            FOR JSON PATH,ROOT('data')
                        )Cad2DFile
                        ,e.Cad2DFileAbsolutePath
                        ,(
                            SELECT  xb.FileId, (xb.[FileName] + xb.FileExtension) FileInfo
                            FROM BAS.[File] xb
                            WHERE e.Cad3DFile=xb.FileId
                            FOR JSON PATH,ROOT('data')
                        )Cad3DFile
                        ,e.Cad3DFileAbsolutePath
                        ,(
                            SELECT  xc.FileId, (xc.[FileName] + xc.FileExtension) FileInfo
                            FROM BAS.[File] xc
                            WHERE e.Pdf2DFile=xc.FileId
                            FOR JSON PATH,ROOT('data')
                        )Pdf2DFile   
                        ,e.Pdf2DFileAbsolutePath                     
                        ,(
                            SELECT  xd.FileId, (xd.[FileName] + xd.FileExtension) FileInfo
                            FROM BAS.[File] xd
                            WHERE e.JmoFile=xd.FileId
                            FOR JSON PATH,ROOT('data')
                        )JmoFile
                        ,e.JmoFileAbsolutePath";
                    sqlQuery.mainTables =
                        @"FROM MES.MoRouting a
                        INNER JOIN MES.RoutingItem b ON a.RoutingItemId = b.RoutingItemId
                        INNER JOIN MES.Routing c ON b.RoutingId = c.RoutingId
                        INNER JOIN MES.ProdMode d ON c.ModeId = d.ModeId
                        LEFT JOIN PDM.RdDesignControl e ON b.ControlId = e.ControlId
                        INNER JOIN MES.ManufactureOrder f ON a.MoId=f.MoId
                        INNER JOIN MES.WipOrder g ON f.WoId=g.WoId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId AND e.ReleasedStatus='Y' ";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);                   
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region//GetMoProcess --取得製令製程 -- Ding 2022.10.14
        public string GetMoProcess(int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.ProcessName,a.MoProcessId,a.ProcessAlias
                             FROM MES.MoProcess a
                             INNER JOIN MES.Process b ON a.ProcessId=b.ProcessId
                             INNER JOIN MES.ManufactureOrder c ON a.MoId=c.MoId
                             INNER JOIN MES.WipOrder d ON c.WoId=d.WoId
                            WHERE  a.MoId= @MoId                            
                    ";
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

        #region//GetJigMoProcess --取治具製令製程 -- Ding 2022.10.14
        public string GetJigMoProcess(string JigNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.MoProcessId,c.ProcessAlias
                            FROM MES.JigBarcode a
                            INNER JOIN MES.Jig b ON a.JigId=b.JigId
                            INNER JOIN MES.MoProcess c ON a.MoProcessId=c.MoProcessId
                            WHERE  b.JigNo= @JigNo                            
                    ";
                    dynamicParameters.Add("JigNo", JigNo);
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

        #region//GetCncProgram --取得編程 -- Ding 2022.10.14
        public string GetCncProgram(int MoProcessId,int CncProgId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.CncProgId,a.CncProgName,a.CncProgApi,a.CncProgNo
                             FROM MES.CncProgram a
                             LEFT JOIN MES.Process b ON a.ProcessId=b.ProcessId
                             INNER JOIN MES.MoProcess c ON b.ProcessId=c.ProcessId
                            WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoProcessId", @" AND c.MoProcessId= @MoProcessId", MoProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CncProgId", @" AND a.CncProgId= @CncProgId", CncProgId);
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

        #region//GetProcessMachine --取得機台 -- Ding 2022.10.14
        public string GetProcessMachine(int MoProcessId,int MachineId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT b.MachineId,b.MachineName,b.MachineDesc,b.CncMachineNo
                             FROM MES.ProcessMachine a
                             INNER JOIN MES.Machine b ON a.MachineId=b.MachineId
                             INNER JOIN MES.ProcessParameter c ON a.ParameterId=c.ParameterId
                             INNER JOIN MES.Process d ON c.ProcessId=d.ProcessId
                             INNER JOIN MES.MoProcess e ON d.ProcessId=e.ProcessId
                            WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoProcessId", @" AND e.MoProcessId= @MoProcessId", MoProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineId", @" AND b.MachineId= @MachineId", MachineId);
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

        #region//GetCncParameter --取得編程參數 --Ding 2022.10.18
        public string GetCncParameter(int CncProgId,string CncParamStatus)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT z.CncProgId,z.CncParamId,z.CncParamNo,z.CncParamName,z.CncParamType,z.CncParamUom
                    ,z.CncParamLevel,z.CncParamValue,z.CncParamStatus
                    FROM (
                        SELECT x.CncProgId,x.CncParamId,x.CncParamNo,x.CncParamName,x.CncParamType,x.CncParamUom
                        ,x.CncParamLevel,x.CncParamValue,x.CncParamStatus
                        FROM (
                                SELECT b.CncProgId, a.CncParamId, a.CncParamNo, a.CncParamName, a.CncParamType, a.CncParamUom
                                , a.CncParamLevel, a.CncParamValue,a.CncParamStatus
                                FROM MES.CncParameter a
                                    INNER JOIN MES.CncProgram b ON a.CncProgId=b.CncProgId
                                WHERE CncParamType ='Text'
                            UNION ALL
                                SELECT DISTINCT b.CncProgId, a.CncParamId, a.CncParamNo, a.CncParamName, a.CncParamType, a.CncParamUom, a.CncParamLevel
                                , (    
                                    SELECT xa.CncParamId, xa.ShowOption, xa.OptionValue
                                    FROM MES.CncParameterOption xa
                                    WHERE a.CncParamId=xa.CncParamId
                                    FOR JSON PATH,ROOT('data')
                                ) CncParamValue,a.CncParamStatus
                                FROM MES.CncParameter a
                                    LEFT JOIN MES.CncProgram b ON a.CncProgId=b.CncProgId
                                    LEFT JOIN MES.CncParameterOption c ON a.CncParamId=c.CncParamId
                                WHERE CncParamType ='Combobox'
                            UNION ALL
                                SELECT b.CncProgId, a.CncParamId, a.CncParamNo, a.CncParamName, a.CncParamType, a.CncParamUom
                                , a.CncParamLevel, a.CncParamValue,a.CncParamStatus
                                FROM MES.CncParameter a
                                    INNER JOIN MES.CncProgram b ON a.CncProgId=b.CncProgId
                                WHERE CncParamType ='File'
                        )x
                    )z
                    WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CncProgId", @" AND z.CncProgId= @CncProgId", CncProgId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CncParamStatus", @" AND z.CncParamStatus= @CncParamStatus", CncParamStatus);
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

        #region//GetCncParameterOption --取得參數選項 --Ding 2022.10.18
        public string GetCncParameterOption(int CncParamId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.ShowOption,a.OptionValue,a.CncParamId
                            FROM MES.CncParameterOption a
                            WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CncParamId", @" AND a.CncParamId= @CncParamId", CncParamId);
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

        #region//GetToolMachineSetting --取得機台所需刀具參數項目設定
        public string GetToolMachineSetting(int ToolId, string ToolNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.ToolId, a.ToolNo
                            , d.ToolClassId
                            , (
                                SELECT e.ToolClassIdSettingId, e.ParameterKey, e.ParameterName
                                FROM MES.ToolClassIdSetting e
                                WHERE e.ToolClassId = d.ToolClassId
                                FOR JSON PATH, ROOT('data')
                            ) ToolClassIdSetting
                            FROM MES.Tool a
                            INNER JOIN MES.ToolModel b ON a.ToolModelId = b.ToolModelId
                            INNER JOIN MES.ToolCategory c ON b.ToolCategoryId = c.ToolCategoryId
                            INNER JOIN MES.ToolClass d ON c.ToolClassId = d.ToolClassId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ToolId", @" AND a.ToolId= @ToolId", ToolId);                   
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ToolNo", @" AND a.ToolNo= @ToolNo", ToolNo);                   
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

        #region//GetProcessMachineToolCount --取得製程機台所綁工具上限數
        public string GetProcessMachineToolCount(int MachineId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.ToolCount
                            FROM MES.ProcessMachine a
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
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

        #region//GetToolMachine --取得機台上工具綁定紀錄
        public string GetToolMachine(int MachineId,int ProcessId,int ModeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT c.ToolModelName+'<第'+CONVERT(varchar,a.MachineLocation)+'刀座>' AS ToolModelName,
                                b.ToolNo, b.ToolId, c.ToolModelId, d.ToolCategoryName, a.MachineLocation
                                                        , e.ToolCount, a.ToolMachineId,f.ProcessId,f.ModeId
                            FROM MES.ToolMachine a
                                INNER JOIN MES.Tool b ON a.ToolId=b.ToolId
                                INNER JOIN MES.ToolModel c ON b.ToolModelId=c.ToolModelId
                                INNER JOIN MES.ToolCategory d ON c.ToolCategoryId=d.ToolCategoryId
                                INNER JOIN MES.ProcessMachine e ON a.MachineId = e.MachineId
                                INNER JOIN MES.ProcessParameter f ON e.ParameterId=f.ParameterId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineId", @" AND a.MachineId= @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessId", @" AND f.ProcessId= @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ModeId", @" AND f.ModeId= @ModeId", ModeId);
                    sql += @" ORDER BY a.MachineLocation";
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

        #region//GetToolMachineParameter --取得機台上工具的設定值
        public string GetToolMachineParameter(int ToolMachineId,int ToolId,int ProcessId,int ModeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT x.ToolId, x.MachineId, x.ToolSpecName, x.ParameterKey , x.ParameterValue,x.MachineLocation, x.ToolNo, x.ToolCount, x.ToolClassIdSettingId,x.ProcessId,x.ModeId
                            FROM (
                                    SELECT DISTINCT a.ToolId, a.MachineId, d.ToolSpecName, d.ToolSpecNo AS ParameterKey , c.ToolSpecValue AS ParameterValue,a.MachineLocation
                                    , e.ToolNo, f.ToolCount, a.ToolMachineId, -1 ToolClassIdSettingId,g.ProcessId,g.ModeId
                                    FROM MES.ToolMachine a
                                    LEFT JOIN MES.ToolSpecLog c ON a.ToolId=c.ToolId
                                    LEFT JOIN MES.ToolSpec d ON c.ToolSpecId=d.ToolSpecId
                                    INNER JOIN MES.Tool e ON a.ToolId = e.ToolId
                                    INNER JOIN MES.ProcessMachine f ON a.MachineId = f.MachineId
                                    INNER JOIN MES.ProcessParameter g ON f.ParameterId=g.ParameterId
                                UNION ALL
                                    SELECT DISTINCT a.ToolId, a.MachineId, b.ParameterName, b.ParameterKey, b.ParameterValue,a.MachineLocation
                                    , c.ToolNo, d.ToolCount, a.ToolMachineId, b.ToolClassIdSettingId,e.ProcessId,e.ModeId
                                    FROM MES.ToolMachine a
                                    LEFT JOIN (
                                    SELECT b.ParameterKey, b.ParameterName, a.ParameterValue, a.ToolMachineId, b.ToolClassIdSettingId
                                    FROM MES.ToolMachineParameter a
                                    INNER JOIN MES.ToolClassIdSetting b ON a.ToolClassIdSettingId=b.ToolClassIdSettingId
                                    ) b ON a.ToolMachineId=b.ToolMachineId
                                    INNER JOIN MES.Tool c ON a.ToolId = c.ToolId
                                    INNER JOIN MES.ProcessMachine d ON a.MachineId = d.MachineId
                                    INNER JOIN MES.ProcessParameter e ON d.ParameterId=e.ParameterId
                            ) x
                            WHERE 1=1
                    ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ToolMachineId", @" AND x.ToolMachineId= @ToolMachineId", ToolMachineId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ToolId", @" AND x.ToolId= @ToolId", ToolId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessId", @" AND x.ProcessId= @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ModeId", @" AND x.ModeId= @ModeId", ModeId);
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

        #region//GetCncProgramJmoFile --取得編程JMO檔案
        public string GetCncProgramJmoFile(int FileId)
        {
            try
            {
                if (FileId <= 0) throw new SystemException("【JMO檔案】查無資料!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT *
                            FROM BAS.[File]
                            WHERE 1=1
                    ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FileId", @" AND FileId= @FileId", FileId);
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

        #region//GetCncProgramRequest --取得CncProgramRequest參數
        public string GetCncProgramRequest(int CncProgLogId)
        {
            try
            {
                if (CncProgLogId <= 0) throw new SystemException("【CncProgLogId】查無資料!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT b.CncParamName,b.CncParamNo,a.RequestValues
                             FROM MES.CncProgramRequest a
                             LEFT JOIN MES.CncParameter b ON a.CncParamId=b.CncParamId
                            WHERE 1=1 AND RequestValues!='-9999'
                            AND a.CncProgLogId= @CncProgLogId
                    ";
                    dynamicParameters.Add("CncProgLogId", CncProgLogId);                    
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

        #region//GetCncProgramResponsest --取得CncProgramResponsest參數
        public string GetCncProgramResponsest(int CncProgLogId,string CncParamNo)
        {
            try
            {
                if (CncProgLogId <= 0) throw new SystemException("【CncProgLogId】不可為空!");
                if (CncParamNo.Equals("")) throw new SystemException("【CncParamNo】不可為空!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT CncProgResponsestId,CncProgLogId,CncParamNo,ResponsestValues,FileId
                            FROM MES.CncProgramResponsest
                            WHERE 1=1
                            AND CncProgLogId= @CncProgLogId
                            AND CncParamNo= @CncParamNo
                    ";
                    dynamicParameters.Add("CncProgLogId", CncProgLogId);
                    dynamicParameters.Add("CncParamNo", CncParamNo);
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

        #region//GetCncProgramWork --取得CncProgramWork結果
        public string GetCncProgramWork(int CncProgramWorkId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CncProgramWorkId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.StartWorkDate,a.EndWorkDate,a.WorkCT
                            ,b.CompanyName,b.ModeName,b.MtlItemName,b.ErpNo,b.ProcessAlias
                            ,c.UserName
                            ,d.MachineDesc";
                    sqlQuery.mainTables =
                        @" FROM MES.CncProgramWork a 
                         INNER JOIN (
                             SELECT z.MoProcessId,w.CompanyName,v.ModeName,u.MtlItemName,y.WoErpPrefix+'-'+y.WoErpNo+'('+CONVERT(varchar(10),x.WoSeq)+')' AS ErpNo,
                                z.ProcessAlias
                             FROM MES.ManufactureOrder x
                             INNER JOIN MES.WipOrder y ON x.WoId=y.WoId
                             INNER JOIN MES.MoProcess z ON z.MoId=x.MoId
                             INNER JOIN BAS.Company w ON y.CompanyId=w.CompanyId
                             INNER JOIN MES.ProdMode v ON x.ModeId=v.ModeId
                             INNER JOIN PDM.MtlItem u ON y.MtlItemId=u.MtlItemId
                             INNER JOIN MES.Process s ON z.ProcessId=s.ProcessId
                         ) b ON b.MoProcessId=a.MoProcessId
                         INNER JOIN BAS.[User] c ON a.CreateBy=c.UserId
                         INNER JOIN MES.Machine d ON a.MachineId=d.MachineId
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CncProgramWorkId", @" AND a.CncProgramWorkId= @CncProgramWorkId", CncProgramWorkId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CncProgramWorkId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetCncProgramWorkFiles --取得報工上傳檔案結果
        public string GetCncProgramWorkFiles(int CncProgramWorkId)
        {
            try
            {
                if (CncProgramWorkId <= 0) throw new SystemException("【CncProgramWorkId】不可為空!");
                
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"  SELECT o.CncProgramWorkFileId,o.CncProgramWorkId,o.FileId,t.[FileName]
                                  FROM MES.CncProgramWorkFiles o
                                  INNER JOIN (
                                        SELECT a.CncProgramWorkId, b.CompanyName, b.ModeName, b.MtlItemName, b.ErpNo, b.ProcessAlias, c.UserName, a.StartWorkDate, a.EndWorkDate, a.WorkCT
                                        FROM MES.CncProgramWork a
                                            INNER JOIN (
                                            SELECT z.MoProcessId, w.CompanyName, v.ModeName, u.MtlItemName, y.WoErpPrefix+'-'+y.WoErpNo+'('+CONVERT(varchar(10),x.WoSeq)+')' AS ErpNo,
                                                z.ProcessAlias
                                            FROM MES.ManufactureOrder x
                                                INNER JOIN MES.WipOrder y ON x.WoId=y.WoId
                                                INNER JOIN MES.MoProcess z ON z.MoId=x.MoId
                                                INNER JOIN BAS.Company w ON y.CompanyId=w.CompanyId
                                                INNER JOIN MES.ProdMode v ON x.ModeId=v.ModeId
                                                INNER JOIN PDM.MtlItem u ON y.MtlItemId=u.MtlItemId
                                                INNER JOIN MES.Process s ON z.ProcessId=s.ProcessId
                                        ) b ON b.MoProcessId=a.MoProcessId
                                        INNER JOIN BAS.[User] c ON a.CreateBy=c.UserId
                                  )p ON o.CncProgramWorkId=p.CncProgramWorkId
                                INNER JOIN BAS.[File] t ON t.FileId=o.FileId 
                              WHERE o.CncProgramWorkId=@CncProgramWorkId
                    ";
                    dynamicParameters.Add("CncProgramWorkId", CncProgramWorkId);                   
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

        #region//GetMoProcessParameter --取得生產模式與製程
        public string GetMoProcessParameter(int MoProcessId)
        {
            try
            {
                if (MoProcessId <= 0) throw new SystemException("【MoProcessId】查無資料!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.*
                         FROM (
                            SELECT x.ModeId,y.ProcessId,y.MoProcessId
                            FROM MES.ManufactureOrder x
                            INNER JOIN MES.MoProcess y ON x.MoId=y.MoId
                         )a
                         INNER JOIN MES.ProcessParameter b ON a.ProcessId=b.ProcessId AND a.ModeId=b.ModeId
                         WHERE a.MoProcessId=@MoProcessId
                    ";
                    dynamicParameters.Add("@MoProcessId", MoProcessId);
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

        #region //GetParameterHistoryData -- 取得編程後參數數值 Ding 20230305
        public string GetParameterHistoryData(int CncProgramWorkId)
        {
            try
            {
                if (CncProgramWorkId <= 0) throw new SystemException("【編程報工】查無資料!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.CncProgLogId,d.CncParamName,c.RequestValues,a.CncProgramWorkId
                         FROM MES.CncProgramWork a
                         INNER JOIN MES.CncProgramLog b ON a.CncProgLogId=b.CncProgLogId
                         INNER JOIN MES.CncProgramRequest c ON c.CncProgLogId=b.CncProgLogId
                         INNER JOIN MES.CncParameter d ON c.CncParamId=d.CncParamId
                         WHERE a.CncProgramWorkId= @CncProgramWorkId
                         ORDER BY d.CncParamLevel
                    ";
                    dynamicParameters.Add("@CncProgramWorkId", CncProgramWorkId);
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

        #region// GetMillToolSpec --取得銑床刀具規格資訊
        public string GetMillToolSpec(string ToolNo)
        {
            try
            {
                if (ToolNo.Length <= 0) throw new SystemException("【工具條碼】查無資料!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT c.ToolNo,b.ToolSpecNo,a.ToolSpecValue
                            FROM MES.ToolSpecLog a
                            INNER JOIN MES.ToolSpec b ON a.ToolSpecId=b.ToolSpecId
                            INNER JOIN MES.Tool c ON a.ToolId=c.ToolId
                            WHERE c.ToolNo= @ToolNo
                    ";
                    dynamicParameters.Add("@ToolNo", ToolNo);
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

        #region// GetMaxToolBlockSeq --取得銑床刀具現況最大刀座編號
        public string GetMaxToolBlockSeq(int MachineId)
        {
            try
            {
                if (MachineId <= 0) throw new SystemException("【機台ID】查無資料!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT ToolBlockSeq
                            FROM MES.MillMachineTool
                            WHERE MachineId=@MachineId
                    ";
                    dynamicParameters.Add("@MachineId", MachineId);
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

        #region//GetMillMachineTool 取得銑床機台刀具資訊
        public string GetMillMachineTool(string OrderBy, int PageIndex, int PageSize,int MachineId, string MachineNo, int MillMachineToolId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MillMachineToolId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  ,a.MachineId,h.ShopId,b.MachineDesc,c.ToolNo,a.ToolBlockSeq,a.ToolBlockName
                            ,a.MillToolB,a.MillToolD,a.MillToolFL,a.MillToolFPR,a.MillToolHD,a.MillToolL,a.MillToolR,a.MillToolRone,a.MillToolRPM
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MillMachineTool a
                            LEFT JOIN MES.Machine b ON a.MachineId=b.MachineId
                            LEFT JOIN MES.Tool c ON a.ToolId=c.ToolId
                            LEFT JOIN MES.ToolModel d ON c.ToolModelId=d.ToolModelId
                            LEFT JOIN MES.ToolCategory e ON e.ToolCategoryId=d.ToolCategoryId
                            LEFT JOIN MES.ToolClass f ON e.ToolClassId=f.ToolClassId
                            LEFT JOIN MES.ToolGroup g ON f.ToolGroupId=g.ToolGroupId
                            LEFT JOIN MES.WorkShop h ON h.ShopId=b.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"";                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MillMachineToolId", @" AND a.MillMachineToolId = @MillMachineToolId", MillMachineToolId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineNo", @" AND b.MachineNo = @MachineNo", MachineNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND b.MachineId = @MachineId", MachineId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MillMachineToolId ASC";
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

        #region//GetMoDwgInfo 依製令取得設計圖
        public string GetMoDwgInfo(string CompanyNo, int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT f.Cad3DFileAbsolutePath
                            FROM MES.ManufactureOrder a
                            INNER JOIN MES.WipOrder b ON a.WoId=b.WoId
                            INNER JOIN PDM.MtlItem c ON b. MtlItemId=c.MtlItemId
                            INNER JOIN MES.MoRouting d ON a.MoId=d.MoId
                            INNER JOIN MES.RoutingItem e on d.RoutingItemId=e.RoutingItemId
                            INNER JOIN PDM.RdDesignControl f ON e.ControlId=f.ControlId
                            WHERE a.MoId=@MoId ";
                    dynamicParameters.Add("MoId", MoId);
                    var Cad3DFileResult = sqlConnection.Query(sql, dynamicParameters);

                    sql = @"SELECT DISTINCT d.JigNo,f.MtlItemNo,f.MtlItemName
                             FROM MES.JigBarcode a
                             INNER JOIN MES.MoProcess b ON a.MoProcessId=b.MoProcessId
                             INNER JOIN MES.ManufactureOrder c ON b.MoId=c.MoId
                             INNER JOIN MES.Jig d ON d.JigId=a.JigId
                             INNER JOIN MES.WipOrder e ON c.WoId=e.WoId
                             INNER JOIN PDM.MtlItem f ON e.MtlItemId=f.MtlItemId
                             WHERE b.MoId=@MoId ";
                    dynamicParameters.Add("MoId", MoId);
                    var JigNoResult = sqlConnection.Query(sql, dynamicParameters);
                    if (JigNoResult.Count() == 0) { throw new SystemException("此製令未綁定治具"); }
                    //JObject JigNoJson = JObject.FromObject(new
                    //{
                    //    data = JigNoResult
                    //});
                    //jsonResponse.Merge(JigNoJson);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data1 = Cad3DFileResult,
                        data2= JigNoResult
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

        #region//GetPreData 取得預編檔案路徑
        public string GetPreData(int MillPreLogId,string JigNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MoProcessId,b.DwgPrtData
                             FROM MES.MillPreEditLog a
                             INNER JOIN MES.MillPreEditRequest b ON a.MillPreLogId=b.MillPreLogId
                            WHERE 1=1 ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MillPreLogId", @" AND a.MillPreLogId= @MillPreLogId", MillPreLogId);
                    var DwgPrtDataResult = sqlConnection.Query(sql, dynamicParameters);
                    if (DwgPrtDataResult.Count() != 1) { throw new SystemException("該【預編檔案路徑】不存在"); }
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = DwgPrtDataResult
                    });
                    #endregion

                    int MoProcessId = -1;
                    foreach (var item in DwgPrtDataResult) {
                        MoProcessId = item.MoProcessId;
                    }

                    sql = @"SELECT COUNT(*) JigQty
                             FROM MES.JigBarcode a
                             INNER JOIN MES.Jig b ON a.JigId=b.JigId
                            WHERE 1=1 AND a.WorkingFlag='N' ";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoProcessId", @" AND a.MoProcessId= @MoProcessId", MoProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "JigNo", @" AND b.JigNo= @JigNo", JigNo);
                    var JigQtyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (JigQtyResult.Count() != 1) { throw new SystemException("此治具尚未綁定製令"); }
                    JObject JigQtyJson = JObject.FromObject(new
                    {
                        data = JigQtyResult
                    });
                    jsonResponse.Merge(JigQtyJson);                   
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

        #region//GetMillProgramWork 取得銑床報工資訊
        public string GetMillProgramWork(string OrderBy, int PageIndex, int PageSize
            , int MoId,string WoErpFullNo, int MoProcessId, int MachineId, string StartDateTime, string EndDateTime)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MillProgramWorkId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  ,f.MtlItemNo,f.MtlItemName,(e.WoErpPrefix + '-' + e.WoErpNo) WoErpFullNo,d.MachineId
                            ,c.WoSeq,b.ProcessAlias,d.MachineDesc,h.JigNo,h.JigName,a.StartWorkDate,a.EndWorkDate";
                    sqlQuery.mainTables =
                        @"FROM MES.MillProgramWork a
                            LEFT JOIN MES.MoProcess b ON a.MoProcessId=b.MoProcessId
                            LEFT JOIN MES.ManufactureOrder c ON b.MoId=c.MoId
                            LEFT JOIN MES.Machine d ON a.MachineId=d.MachineId
                            LEFT JOIN MES.WipOrder e ON c.WoId=e.WoId
                            LEFT JOIN PDM.MtlItem f ON e.MtlItemId=f.MtlItemId
                            LEFT JOIN MES.MillProgramLog g ON a.MillProgLogId=g.MillProgLogId
                            LEFT JOIN MES.Jig h ON g.JigId=h.JigId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND (a.WoErpPrefix + '-' + a.WoErpNo) LIKE '%' + @WoErpFullNo + '%'", WoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND d.MachineId = @MachineId", MachineId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MillProgramWorkId DESC";
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

        #region//GetMillData 銑床檔案下載
        public string GetMillData(string OrderBy, int PageIndex, int PageSize
                    , string DataType, int MillProgramWorkId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    switch (DataType) {
                        case "Dwg":
                            sqlQuery.mainKey = "d.MillPreRequestId";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @",d.DwgPrtData AS Data";
                            sqlQuery.mainTables =
                                @"FROM MES.MillProgramWork a
                                    INNER JOIN MES.MillProgramLog b ON a.MillProgLogId=b.MillProgLogId
                                    INNER JOIN MES.MillPreEditLog c ON b.MillPreLogId=c.MillPreLogId
                                    INNER JOIN MES.MillPreEditRequest d ON c.MillPreLogId=d.MillPreLogId";
                            string queryTable = "";
                            sqlQuery.auxTables = queryTable;
                            string queryCondition = @"";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MillProgramWorkId", @" AND a.MillProgramWorkId = @MillProgramWorkId", MillProgramWorkId);
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MillProgramWorkId ASC";
                            sqlQuery.pageIndex = PageIndex;
                            sqlQuery.pageSize = PageSize;
                            break;
                        case "PrePrt":
                            sqlQuery.mainKey = "e.MillPreResponsestId";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @",e.MillPrePrtData  AS Data";
                            sqlQuery.mainTables =
                                @"FROM MES.MillProgramWork a
                                INNER JOIN MES.MillProgramLog b ON a.MillProgLogId=b.MillProgLogId
                                INNER JOIN MES.MillPreEditLog c ON b.MillPreLogId=c.MillPreLogId
                                INNER JOIN MES.MillPreEditRequest d ON c.MillPreLogId=d.MillPreLogId
                                INNER JOIN MES.MillPreEditResponsest e ON c.MillPreLogId=e.MillPreLogId";
                            queryTable = "";
                            sqlQuery.auxTables = queryTable;
                            queryCondition = @"";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MillProgramWorkId", @" AND a.MillProgramWorkId = @MillProgramWorkId", MillProgramWorkId);
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MillProgramWorkId ASC";
                            sqlQuery.pageIndex = PageIndex;
                            sqlQuery.pageSize = PageSize;
                            break;
                        case "Prt":
                            sqlQuery.mainKey = "c.MillProgResponsestId";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @",c.MillPrtData  AS Data";
                            sqlQuery.mainTables =
                                @"FROM MES.MillProgramWork a
                                    INNER JOIN MES.MillProgramLog b ON a.MillProgLogId=b.MillProgLogId                                    
                                    INNER JOIN MES.MillProgramResponsest c ON b.MillProgLogId=c.MillProgLogId";
                            queryTable = "";
                            sqlQuery.auxTables = queryTable;
                            queryCondition = @"";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MillProgramWorkId", @" AND a.MillProgramWorkId = @MillProgramWorkId", MillProgramWorkId);
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MillProgramWorkId ASC";
                            sqlQuery.pageIndex = PageIndex;
                            sqlQuery.pageSize = PageSize;
                            break;
                    }
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

        #region//Add

        #region //AddCncProgram -- 編程程式基本資料 新增 --Ding 2022.10.04
        public string AddCncProgram(int ProcessId, string CncProgName, string CncProgApi)
        {
            try
            {
                if (ProcessId <= 0) throw new SystemException("【生產製程】不能為空!");
                if (CncProgName.Length <= 0) throw new SystemException("【CNC編程名稱】不能為空!");
                if (CncProgApi.Length <= 0) throw new SystemException("【CNC編程API名稱】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【CNC編程名稱】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CncProgName 
                                FROM MES.CncProgram a                                
                                WHERE a.CncProgName =@CncProgName
                                AND a.ProcessId =@ProcessId";
                        dynamicParameters.Add("CncProgName", CncProgName);
                        dynamicParameters.Add("ProcessId", ProcessId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該【CNC編程名稱】已存在");
                        #endregion

                        #region //判斷【CNC編程API名稱】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CncProgApi 
                                FROM MES.CncProgram a                                
                                WHERE a.CncProgApi =@CncProgApi
                                AND a.ProcessId =@ProcessId";
                        dynamicParameters.Add("CncProgApi", CncProgApi);
                        dynamicParameters.Add("ProcessId", ProcessId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該【CNC編程API名稱】已存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.CncProgram (ProcessId, CncProgName, CncProgApi, 
                                [Status], CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.CncProgId
                                VALUES (@ProcessId, @CncProgName, @CncProgApi, 
                                @Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessId,
                                CncProgName,
                                CncProgApi,
                                Status = "Y",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddCncProgramFile -- 編程程式檔案基本資料 新增 --Ding 2022.10.04
        public string AddCncProgramFile(int CncProgId, string FileDesc, int FileId)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region //AddCncProgram -- 編程程式參數基本資料 新增 --Ding 2022.10.04
        public string AddCncParameter(int CncProgId, string CncParamNo, string CncParamName, string CncParamType, string CncParamOption
            , string CncParamUomId, string CncParamValue, string CncParamStatus)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region //AddCncProgramKnife -- 編程程式刀具基本資料 新增 --Ding 2022.10.04
        public string AddCncProgramKnife(int CncProgId, string KnifeName, string KnifeKey)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region//AddCncProgramLog --編程執行紀錄資料 新增 --Ding 2022.10.17
        public string AddCncProgramLog(int CncProgId, int MoProcessId,int MachineId)
        {
            try
            {
                if (MoProcessId <= 0) throw new SystemException("【製令製程ID】不能為空!");
                if (CncProgId <= 0) throw new SystemException("【CNC編程ID】不能為空!");
                if (MachineId <= 0) throw new SystemException("【機台ID】不能為空!");


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【製令製程ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoProcessId 
                                FROM MES.MoProcess a                                
                                WHERE a.MoProcessId =@MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【製令製程ID】不存在");
                        #endregion

                        #region //判斷【CNC編程ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CncProgId 
                                FROM MES.CncProgram a                              
                                WHERE a.CncProgId =@CncProgId";
                        dynamicParameters.Add("CncProgId", CncProgId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【CNC編程ID】不存在");
                        #endregion

                        #region //判斷【機台ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MachineId 
                                FROM MES.Machine a                              
                                WHERE a.MachineId =@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("該【機台ID】不存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO  MES.CncProgramLog (CncProgId, MoProcessId,MachineId
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.CncProgLogId
                                VALUES (@CncProgId, @MoProcessId, @MachineId
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgId,
                                MoProcessId,
                                MachineId,
                                Status = "A",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddToolMachineParameter --機台所需刀具參數 新增 --Ding 2022.10.17
        public string AddToolMachineParameter(int MachineId, string ToolNo, int MachineLocation, int ToolClassIdSettingId, string ParameterValue, int AddCount, int ToolMachineId)
        {
            try
            {
                if (MachineId <= 0) throw new SystemException("【機台ID】不能為空!");
                if (ToolNo == "") throw new SystemException("【工具No】不能為空!");
                if (MachineLocation <= 0) throw new SystemException("【機台座台】不能為空!");
                if (ToolClassIdSettingId <= 0) throw new SystemException("【工具機台設定ID】不能為空!");
                if (ParameterValue == "") throw new SystemException("【工具機台設定Value】不能為空!");
                int ToolId=-1;
                int rowsAffected = 0;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【機台ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MachineId 
                                FROM MES.Machine a                                
                                WHERE a.MachineId =@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【機台ID】不存在");
                        #endregion

                        #region //判斷【工具ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToolId 
                                FROM MES.Tool a                              
                                WHERE a.ToolNo =@ToolNo";
                        dynamicParameters.Add("ToolNo", ToolNo);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【工具NO】不存在");
                        foreach (var item in result2)
                        {
                            ToolId = item.ToolId;
                        }
                        if (ToolId == -1) throw new SystemException("【工具ID】不能為-1");
                        #endregion

                        #region //判斷【工具機台設定ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToolClassIdSettingId 
                                FROM MES.ToolClassIdSetting a                              
                                WHERE a.ToolClassIdSettingId =@ToolClassIdSettingId";
                        dynamicParameters.Add("ToolClassIdSettingId", ToolClassIdSettingId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("該【工具機台設定ID】不存在");
                        #endregion

                        #region//查詢此機台的刀座是否已經有刀具在上面
                        if (AddCount == 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ToolMachineId 
                                    FROM MES.ToolMachine a                              
                                    WHERE a.MachineId =@MachineId
                                    AND a.MachineLocation=@MachineLocation";
                            dynamicParameters.Add("MachineId", MachineId);
                            dynamicParameters.Add("MachineLocation", MachineLocation);
                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() > 0) throw new SystemException("該【刀座】已綁定刀具");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ToolMachineId 
                                    FROM MES.ToolMachine a                              
                                    WHERE a.ToolId =@ToolId";
                            dynamicParameters.Add("ToolId", ToolId);
                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() > 0) throw new SystemException("該【刀具】已被綁定");

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolMachine (MachineId,ToolId,MachineLocation
                                    ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.ToolMachineId
                                    VALUES (@MachineId, @ToolId, @MachineLocation
                                    ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MachineId,
                                    ToolId,
                                    MachineLocation,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();

                            foreach (var item in insertResult)
                            {
                                ToolMachineId = item.ToolMachineId;
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
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolMachineParameter (ToolMachineId, ToolClassIdSettingId, ParameterValue
                                , [Status], CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolMachineParameterId, INSERTED.ToolMachineId
                                VALUES (@ToolMachineId, @ToolClassIdSettingId, @ParameterValue
                                , @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolMachineId,
                                ToolClassIdSettingId,
                                ParameterValue,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult2.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult2
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

        #region//AddCncProgramRequest --編程需求參數數值 新增 --Ding 2022.11.7
        public string AddCncProgramRequest(int CncProgLogId, string RequestKey, string RequestValues)
        {
            try
            {
                if (CncProgLogId <= 0) throw new SystemException("【編程Log ID】不能為空!");
                if (RequestKey.Equals("")) throw new SystemException("【編程參數】不能為空!");
                if (RequestValues.Equals("")) throw new SystemException("【編程參數數值】不能為空!");                

                using (TransactionScope transactionScope = new TransactionScope())
                {                   
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【編程Log ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CncProgId
                                FROM MES.CncProgramLog                       
                                WHERE CncProgLogId =@CncProgLogId";
                        dynamicParameters.Add("CncProgLogId", CncProgLogId);                        
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【編程Log ID】不存在");
                        int CncProgId = -1;
                        foreach (var item in result)
                        {
                            CncProgId = item.CncProgId;
                        }
                        #endregion

                        #region //判斷【編程參數】是否存在
                        dynamicParameters = new DynamicParameters();
                        int CncParamId = -1;
                        sql = @"SELECT CncParamId,CncParamNo
                                FROM MES.CncParameter                             
                                WHERE CncProgId= @CncProgId
                                AND CncParamNo =@CncParamNo";
                        dynamicParameters.Add("CncProgId", CncProgId);
                        dynamicParameters.Add("CncParamNo", RequestKey);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() != 1) {
                            throw new SystemException("該【編程參數】不存在");
                        }else{
                            foreach (var item in result2)
                            {
                                CncParamId = item.CncParamId;
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO  MES.CncProgramRequest (CncProgLogId,CncParamId,RequestValues
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@CncProgLogId, @CncParamId, @RequestValues
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgLogId,
                                CncParamId,
                                RequestValues,
                                Status = "A",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddCncProgramResponsest --新增編程回傳參數
        public string AddCncProgramResponsest(int CncProgLogId, string CncParamNo, string ResponsestValues,int FileId)
        {
            try
            {
                if (CncProgLogId <= 0) throw new SystemException("【回傳編程Log ID】不能為空!");
                if (CncParamNo.Equals("")) throw new SystemException("【回傳編程參數】不能為空!");
                if (ResponsestValues.Equals("")) throw new SystemException("【回傳編程參數數值】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【編程Log ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CncProgLogId
                                FROM MES.CncProgramLog                       
                                WHERE CncProgLogId =@CncProgLogId";
                        dynamicParameters.Add("CncProgLogId", CncProgLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【編程Log ID】不存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.CncProgramResponsest (CncProgLogId,CncParamNo,ResponsestValues,FileId
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@CncProgLogId, @CncParamNo, @ResponsestValues,@FileId
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgLogId,
                                CncParamNo,
                                ResponsestValues,
                                FileId,
                                Status = "A",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddCncResponsestFile --新增編程回傳檔案 --新增編程回傳檔案--暫時用不到了
        //public string AddCncResponsestFile(int Company, string FileName, string FileContent, string FileExtension, string FileSize, string ClientIP, string Source)
        //{
        //    try
        //    {
        //        if (Company <= 0) throw new SystemException("【公司ID】不能為空!");
        //        if (FileName.Equals("")) throw new SystemException("【檔名】不能為空!");
        //        if (FileContent.Equals("")) throw new SystemException("【檔案內容】不能為空!");
        //        if (FileExtension.Equals("")) throw new SystemException("【副檔名】不能為空!");
        //        if (FileSize.Equals("")) throw new SystemException("【檔案大小】不能為空!");
        //        if (ClientIP.Equals("")) throw new SystemException("【來源IP】不能為空!");
        //        if (Source.Equals("")) throw new SystemException("【來源】不能為空!");

        //        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
        //        {
        //            using (TransactionScope transactionScope = new TransactionScope())
        //            {
        //                DynamicParameters dynamicParameters = new DynamicParameters();

        //                #region //判斷【公司ID】是否存在
        //                dynamicParameters = new DynamicParameters();
        //                sql = @"SELECT CompanyId
        //                        FROM BAS.Company                       
        //                        WHERE CompanyId =@CompanyId";
        //                dynamicParameters.Add("CompanyId", Company);
        //                var result = sqlConnection.Query(sql, dynamicParameters);
        //                if (result.Count() <= 0) throw new SystemException("該【公司ID】不存在");
        //                #endregion

        //                dynamicParameters = new DynamicParameters();
        //                sql = @"INSERT INTO  BAS.[File] (CompanyId,[FileName],FileContent,FileExtension,FileSize,ClientIP,Source
        //                        ,[DeleteStatus],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
        //                        OUTPUT INSERTED.FileId
        //                        VALUES (@Company, @FileName, @FileContent,@FileExtension,@FileSize,@ClientIP,@Source
        //                        ,@DeleteStatus,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
        //                dynamicParameters.AddDynamicParams(
        //                    new
        //                    {
        //                        Company,
        //                        FileName,
        //                        FileContent,
        //                        FileExtension,
        //                        FileSize,
        //                        ClientIP,
        //                        Source,
        //                        DeleteStatus = "N",
        //                        CreateDate,
        //                        LastModifiedDate,
        //                        CreateBy,
        //                        LastModifiedBy
        //                    });
        //                var insertResult = sqlConnection.Query(sql, dynamicParameters);
        //                int rowsAffected = insertResult.Count();

        //                #region //Response
        //                jsonResponse = JObject.FromObject(new
        //                {
        //                    status = "success",
        //                    msg = "(" + rowsAffected + " rows affected)",
        //                    data = insertResult
        //                });
        //                #endregion

        //                transactionScope.Complete();
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "errorForDA",
        //            msg = e.Message
        //        });
        //        #endregion

        //        logger.Error(e.Message);
        //    }

        //    return jsonResponse.ToString();
        //}
        #endregion

        #region //AddCncResponsestFile-- 新增python回傳檔案 -- Ding 2022-11-09
        public string AddCncResponsestFile(byte[] FileContent, string FileName, string FileExtension, int FileSize, string ClientIP, string Source)
        {
            try
            {
                if (FileContent.Length <= 0) throw new SystemException("【檔案Binary】不能為空!");
                if (FileName.Length <= 0) throw new SystemException("【檔案名稱】不能為空!");
                if (FileExtension.Length <= 0) throw new SystemException("【檔案副檔名】不能為空!");
                if (FileSize <= 0) throw new SystemException("【檔案大小】不能為空!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize, ClientIP, Source, DeleteStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FileId
                                VALUES (@CompanyId, @FileName, CONVERT(varbinary(max),@FileContent), @FileExtension, @FileSize, @ClientIP, @Source, @DeleteStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                FileName,
                                FileContent,
                                FileExtension,
                                FileSize,
                                ClientIP,
                                Source,
                                DeleteStatus = "N",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddCncDirectWork --直接編程報工 新增
        public string AddCncDirectWork(int CncProgLogId, int MoProcessId, int MachineId, string StartWorkDate, string EndWorkDate, int WorkCT,
            string FileArray)
        {
            try
            {
                if (CncProgLogId <= 0) throw new SystemException("【CNC編程LogID】不能為空!");
                if (MoProcessId <= 0) throw new SystemException("【製令製程ID】不能為空!");
                if (MachineId <= 0) throw new SystemException("【機台ID】不能為空!");
                if (StartWorkDate.Equals("")) throw new SystemException("【開工時間】不能為空!");
                if (EndWorkDate.Equals("")) throw new SystemException("【完工時間】不能為空!");
                if (WorkCT <= 0) throw new SystemException("【工時】不能為空!");
                if (FileArray.Equals("")) throw new SystemException("【上傳檔案】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【CNC編程LogID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CncProgLogId 
                                FROM MES.CncProgramLog a                                
                                WHERE a.CncProgLogId =@CncProgLogId";
                        dynamicParameters.Add("CncProgLogId", CncProgLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【CNC編程LogID】不存在");
                        #endregion

                        #region //判斷【製令製程ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoProcessId 
                                FROM MES.MoProcess a                              
                                WHERE a.MoProcessId =@MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【製令製程ID】不存在");
                        #endregion

                        #region //判斷【機台ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MachineId 
                                FROM MES.Machine a                              
                                WHERE a.MachineId =@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("該【機台ID】不存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.CncProgramWork (CncProgLogId,MoProcessId,MachineId,StartWorkDate,EndWorkDate,WorkCT
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.CncProgramWorkId
                                VALUES (@CncProgLogId,@MoProcessId,@MachineId,@StartWorkDate,@EndWorkDate,@WorkCT
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgLogId,
                                MoProcessId,
                                MachineId,
                                StartWorkDate,
                                EndWorkDate,
                                WorkCT,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int CncProgramWorkId = -1;
                        foreach (var item in insertResult)
                        {
                            CncProgramWorkId = item.CncProgramWorkId;
                        }

                        List<string> KeyList = new List<string>();
                        JObject FileArrayJson = JObject.Parse(FileArray);
                        foreach (var x in FileArrayJson)
                        {
                            KeyList.Add(x.Key);
                        }

                        foreach (string item in KeyList) {
                            string FileIdArray = FileArrayJson[item][0]["FileId"].ToString();
                            string FileType = FileArrayJson[item][0]["FileType"].ToString();
                            string NcName = FileArrayJson[item][0]["NcName"].ToString();
                            string[] FileIdItem = FileIdArray.Split(',');
                            foreach (var id in FileIdItem)
                            {
                                #region //判斷【File ID】是否存在
                                int FileId = int.Parse(id);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a                              
                                WHERE a.FileId =@FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var result4 = sqlConnection.Query(sql, dynamicParameters);
                                if (result4.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.CncProgramWorkFiles (CncProgramWorkId,FileId,NcName,FileType
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@CncProgramWorkId,@FileId,@NcName,@FileType
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CncProgramWorkId,
                                        FileId,
                                        NcName,
                                        FileType,
                                        Status = "A",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);
                            }
                            
                        }
                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddCncProgramWork --編程報工 新增
        public string AddCncProgramWork(int CncProgLogId, int MoProcessId, int MachineId, string StartWorkDate, string EndWorkDate, int WorkCT,
            string FileArray)
        {
            try
            {
                if (CncProgLogId <= 0) throw new SystemException("【CNC編程LogID】不能為空!");
                if (MoProcessId <= 0) throw new SystemException("【製令製程ID】不能為空!");
                if (MachineId <= 0) throw new SystemException("【機台ID】不能為空!");
                if (StartWorkDate.Equals("")) throw new SystemException("【開工時間】不能為空!");
                if (EndWorkDate.Equals("")) throw new SystemException("【完工時間】不能為空!");
                if (WorkCT <= 0) throw new SystemException("【工時】不能為空!");
                if (FileArray.Equals("")) throw new SystemException("【上傳檔案】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【CNC編程LogID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CncProgLogId 
                                FROM MES.CncProgramLog a                                
                                WHERE a.CncProgLogId =@CncProgLogId";
                        dynamicParameters.Add("CncProgLogId", CncProgLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【CNC編程LogID】不存在");
                        #endregion

                        #region //判斷【製令製程ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoProcessId 
                                FROM MES.MoProcess a                              
                                WHERE a.MoProcessId =@MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【製令製程ID】不存在");
                        #endregion

                        #region //判斷【機台ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MachineId 
                                FROM MES.Machine a                              
                                WHERE a.MachineId =@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("該【機台ID】不存在");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.CncProgramWork (CncProgLogId,MoProcessId,MachineId,StartWorkDate,EndWorkDate,WorkCT
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.CncProgramWorkId
                                VALUES (@CncProgLogId,@MoProcessId,@MachineId,@StartWorkDate,@EndWorkDate,@WorkCT
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgLogId,
                                MoProcessId,
                                MachineId,
                                StartWorkDate,
                                EndWorkDate,
                                WorkCT,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int CncProgramWorkId = -1;
                        foreach (var item in insertResult)
                        {
                            CncProgramWorkId = item.CncProgramWorkId;
                        }

                        List<string> KeyList = new List<string>();
                        JObject FileArrayJson = JObject.Parse(FileArray);
                        foreach (var x in FileArrayJson)
                        {
                            KeyList.Add(x.Key);
                        }

                        int ItemCount = KeyList.Count();
                        int totalCount = -1;
                        for (int j=0;j< ItemCount;j++) {
                            totalCount = FileArrayJson[KeyList[j]].Count();
                            for (int i = 0; i < totalCount; i++)
                            {
                                string FileIdArray = FileArrayJson[KeyList[j]][i]["FileId"].ToString();
                                string FileType = FileArrayJson[KeyList[j]][i]["FileType"].ToString();
                                string NcName = FileArrayJson[KeyList[j]][i]["NcName"].ToString();
                                string[] FileIdItem = FileIdArray.Split(',');
                                foreach (var Item in FileIdItem)
                                {
                                    #region //判斷【File ID】是否存在
                                    int FileId = int.Parse(FileIdArray);
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.FileId 
                                    FROM BAS.[File] a                              
                                    WHERE a.FileId =@FileId";
                                    dynamicParameters.Add("FileId", FileId);
                                    var result4 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result4.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                    #endregion

                                    if (NcName.Contains(":") == true)
                                    {
                                        string[] x = NcName.Split(':');
                                        NcName = x[1].ToString();
                                    }
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.CncProgramWorkFiles (CncProgramWorkId,FileId,NcName,FileType
                                    ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                    VALUES (@CncProgramWorkId,@FileId,@NcName,@FileType
                                    ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CncProgramWorkId,
                                            FileId,
                                            NcName,
                                            FileType,
                                            Status = "A",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);
                                }
                            }
                        }
                        
                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddMillMachineTool --銑床機台刀具 新增
        public string AddMillMachineTool(string ToolNo, int MachineId, string ToolBlockSeq, string ToolBlockName
            , string MillToolD, string MillToolRone, string MillToolB, string MillToolFL, string MillToolL, string MillToolHD
            , string MillToolR, string MillToolRPM, string MillToolFPR)
        {
            try
            {
                if (ToolNo.Equals("") ) throw new SystemException("【工具編號】不能為空!");
                if (MachineId <= 0) throw new SystemException("【機台ID】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【工具ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ToolId
                                FROM MES.Tool                               
                                WHERE ToolNo =@ToolNo";
                        dynamicParameters.Add("ToolNo", ToolNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【工具ID】不存在");
                        int ToolId= sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolId;
                        #endregion

                        #region //判斷【機台ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine                              
                                WHERE MachineId =@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【機台ID】不存在");
                        #endregion

                        #region//同機台不可出現重複的刀座
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MillMachineTool                              
                                WHERE MachineId =@MachineId
                                AND ToolBlockSeq=@ToolBlockSeq";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("ToolBlockSeq", ToolBlockSeq);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【該機台不可出現重複的刀座】");

                        #endregion

                        #region//已經維護的刀具，不可以同時出現在2台機台上
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.MachineDesc
                                FROM MES.MillMachineTool a
                                INNER JOIN MES.Machine b ON a.MachineId=b.MachineId                            
                                WHERE a.ToolId =@ToolId
                                AND a.MachineId=@MachineId";
                        dynamicParameters.Add("ToolId", ToolId);
                        dynamicParameters.Add("MachineId", MachineId);
                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 1)
                        {
                            foreach (var item in result4)
                            {
                                throw new SystemException("該刀具目前架設在【" + item.MachineDesc + "】，不可新增");
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillMachineTool (ToolId,MachineId,ToolBlockSeq,ToolBlockName
                                ,MillToolD,MillToolRone,MillToolB,MillToolFL,MillToolL,MillToolHD,MillToolR,MillToolRPM,MillToolFPR
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillMachineToolId
                                VALUES (@ToolId,@MachineId,@ToolBlockSeq,@ToolBlockName
                                ,@MillToolD,@MillToolRone,@MillToolB,@MillToolFL,@MillToolL,@MillToolHD,@MillToolR,@MillToolRPM,@MillToolFPR
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolId,
                                MachineId,
                                ToolBlockSeq,
                                ToolBlockName,
                                MillToolD,
                                MillToolRone,
                                MillToolB,
                                MillToolFL,
                                MillToolL,
                                MillToolHD,
                                MillToolR,
                                MillToolRPM,
                                MillToolFPR,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床機台刀具】新增失敗!");

                        #region//新增工具異動紀錄
                        string TransactionType = "", ToolLocatorNameNow="", ToolInventoryNameNow="";
                        int ToolLocatorIdNow = 0;

                        #region //判斷【機台工具儲位】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ToolLocatorId
                                FROM MES.MachineToolLocator                               
                                WHERE MachineId = @MachineId";
                        dynamicParameters.Add("MachineId", MachineId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台工具儲位不存在，請重新輸入!");
                        int ToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                        #endregion

                        #region //判斷【工具儲位】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorId = @ToolLocatorId";
                        dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                        string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                        string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                        #endregion

                        #region //判斷工具是否有入庫過
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @ToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                        dynamicParameters.Add("ToolId", ToolId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                            if (TransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                        }
                        else
                        {
                            ToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            ToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            ToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            if (ToolLocatorIdNow == ToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                        }
                        #endregion
                      
                        if (ToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolId,
                                    TransactionType = "Out",
                                    TransactionDate = DateTime.Now,
                                    ToolLocatorId = ToolLocatorIdNow,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult01.Count();
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolId,
                                TransactionType = "In",
                                TransactionDate = DateTime.Now,
                                ToolLocatorId,
                                TraderId = CreateBy,
                                TransactionReason =  CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】" ,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy,
                                ProcessingQty=0
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region //AddMillPreEditLog 新增銑床預編Log資訊 (API)
        public string AddMillPreEditLog(int MoProcessId,string JigNo, string ServerPath, string DwgData)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();   
                        #region //判斷【使用者】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.MoProcessId,c.UserId,b.JigNo,e.CompanyId,e.CompanyNo
                                FROM MES.JigBarcode a
                                INNER JOIN MES.Jig b ON a.JigId=b.JigId
                                INNER JOIN BAS.[User] c ON a.CreateBy=c.UserId
                                INNER JOIN BAS.Department d ON d.DepartmentId=c.DepartmentId
                                INNER JOIN BAS.Company e ON d.CompanyId=e.CompanyId                           
                                WHERE a.MoProcessId =@MoProcessId AND b.JigNo=@JigNo";
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        dynamicParameters.Add("JigNo", JigNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【使用者】不存在");
                        int UserId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).UserId;
                        int UserCompanyId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CompanyId;
                        string CompanyNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CompanyNo;
                        #endregion

                        #region //確認【製令製程】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT c.CompanyId
                                 FROM MES.MoProcess a
                                 INNER JOIN MES.ManufactureOrder b ON a.MoId=b.MoId
                                 INNER JOIN MES.WipOrder c ON b.WoId=c.WoId                             
                                WHERE a.MoProcessId =@MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【製令製程】不存在");                       
                        int MoCompanyId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CompanyId;
                        #endregion

                        if (MoCompanyId!= UserCompanyId) { throw new SystemException("使用者【所屬公司】與製令不同"); }

                        #region//銑床預編Log
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillPreEditLog (MoProcessId
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillPreLogId
                                VALUES (@MoProcessId
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MoProcessId,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床預編Log】新增失敗!");
                        int MillPreLogId = -1;
                        foreach (var item in insertResult)
                        {
                            MillPreLogId = item.MillPreLogId;
                        }

                        #endregion

                        #region//檔案處理  
                        string FolderPath = Path.Combine(ServerPath, CompanyNo, "MillPreEdit", "Request");
                        string ASCIIPath = "";
                        string FilePath = "";
                        ASCIIPath = DwgData;
                        if (DwgData.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                        if (DwgData.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                        if (DwgData.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                        if (DwgData.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                        if (DwgData.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                        if (DwgData.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");

                        if (File.Exists(ASCIIPath))
                        {
                            string FileName = Path.GetFileNameWithoutExtension(DwgData);
                            string FileExtension = Path.GetExtension(DwgData);
                            FilePath = Path.Combine(FolderPath, FileName + "(" + MillPreLogId+")")+ FileExtension;
                            if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                            byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                            File.WriteAllBytes(FilePath, cadFileByte);
                        }
                        else
                        {
                            throw new SystemException("查無設計圖路徑!!");
                        }

                        #endregion

                        #region//銑床預編Request
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillPreEditRequest (MillPreLogId,DwgPrtData
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillPreLogId
                                VALUES (@MillPreLogId,@DwgPrtData
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MillPreLogId,
                                DwgPrtData= FilePath,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected = insertResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床預編Log】新增失敗!");
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddMillPreEditResponsest 新增銑床預編回傳紀錄(API)
        public string AddMillPreEditResponsest(int MillPreLogId,string CompanyNo, string ServerPath,string MillPrePrtData)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //判斷【銑床預編執行紀錄Log】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MillPreLogId,LastModifiedBy
                                 FROM MES.MillPreEditLog                            
                                WHERE MillPreLogId =@MillPreLogId";
                        dynamicParameters.Add("MillPreLogId", MillPreLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【銑床預編執行紀錄Log】不存在");
                        int UserId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).LastModifiedBy;                  
                        #endregion

                        #region//檔案處理  
                        string FolderPath = Path.Combine(ServerPath, CompanyNo, "MillPreEdit", "Responsest");
                        string ASCIIPath = "";
                        string FilePath = "";
                        ASCIIPath = MillPrePrtData;
                        if (MillPrePrtData.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                        if (MillPrePrtData.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                        if (MillPrePrtData.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                        if (MillPrePrtData.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                        if (MillPrePrtData.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                        if (MillPrePrtData.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");

                        if (File.Exists(ASCIIPath))
                        {
                            string FileName = Path.GetFileNameWithoutExtension(MillPrePrtData);
                            string FileExtension = Path.GetExtension(MillPrePrtData);
                            FilePath = Path.Combine(FolderPath, FileName + "(" + MillPreLogId + ")"+ FileExtension);
                            if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                            byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                            File.WriteAllBytes(FilePath, cadFileByte);
                        }
                        else
                        {
                            throw new SystemException("查無預編檔路徑!!");
                        }

                        #endregion

                        #region//銑床預編MillProgramResponsest
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillPreEditResponsest (MillPreLogId,MillPrePrtData
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillPreLogId
                                VALUES (@MillPreLogId,@MillPrePrtData
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MillPreLogId,
                                MillPrePrtData = FilePath,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床預編回傳紀錄】新增失敗!");
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region//AddMillProgramLog 新增銑床編程執行紀錄(API)
        public string AddMillProgramLog(string CompanyNo,string SecretKey, int MillPreLogId,string MachineNo, string MillToolData,string MillPrePrtData,string ServerPath, string JigNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //判斷【銑床預編執行紀錄Log】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MillPreLogId,e.CompanyId,a.CreateBy
                                FROM MES.MillPreEditLog a
                                INNER JOIN MES.MoProcess b ON a.MoProcessId=b.MoProcessId
                                INNER JOIN MES.ManufactureOrder c ON b.MoId=c.MoId
                                INNER JOIN MES.WipOrder d ON d.WoId=c.WoId
                                INNER JOIN BAS.Company e ON d.CompanyId=e.CompanyId                           
                                WHERE a.MillPreLogId =@MillPreLogId";
                        dynamicParameters.Add("MillPreLogId", MillPreLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【銑床預編執行紀錄Log】不存在");
                        int PreCompanyId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CompanyId;
                        int UserId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CreateBy;
                        #endregion

                        #region//判斷【機台】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MachineId
                                FROM MES.Machine                       
                                WHERE MachineNo =@MachineNo";
                        dynamicParameters.Add("MachineNo", MachineNo);
                        var MachineResult = sqlConnection.Query(sql, dynamicParameters);
                        if (MachineResult.Count() != 1) throw new SystemException("該【機台】不存在");
                        int MachineId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MachineId;
                        #endregion

                        #region//判斷治具是否與製令有綁定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT JigId
                                FROM MES.Jig                      
                                WHERE JigNo =@JigNo";
                        dynamicParameters.Add("JigNo", JigNo);
                        var JigResult = sqlConnection.Query(sql, dynamicParameters);
                        if (JigResult.Count() != 1) throw new SystemException("該【治具】不存在");
                        int JigId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).JigId;
                        #endregion

                        #region//銑床預編Log
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillProgramLog (MillPreLogId,MachineId,JigId
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillProgLogId
                                VALUES (@MillPreLogId,@MachineId,@JigId
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MillPreLogId,
                                MachineId,
                                JigId,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var MillProgLogIdResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = MillProgLogIdResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床預編Log】新增失敗!");
                        int MillProgLogId = -1;
                        foreach (var item in MillProgLogIdResult)
                        {
                            MillProgLogId = item.MillProgLogId;
                        }
                        #endregion

                        #region//檔案處理  
                        string FolderPath = Path.Combine(ServerPath, CompanyNo, "MillProgram", "Request");
                        string ASCIIPath = "";
                        string FilePath = "";
                        ASCIIPath = MillPrePrtData;
                        if (MillPrePrtData.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                        if (MillPrePrtData.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                        if (MillPrePrtData.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                        if (MillPrePrtData.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                        if (MillPrePrtData.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                        if (MillPrePrtData.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");

                        if (File.Exists(ASCIIPath))
                        {
                            string FileName = Path.GetFileNameWithoutExtension(MillPrePrtData);
                            string FileExtension = Path.GetExtension(MillPrePrtData);
                            FilePath = Path.Combine(FolderPath, FileName + "(" + MillProgLogId + ")"+ FileExtension);
                            if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                            byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                            File.WriteAllBytes(FilePath, cadFileByte);
                        }
                        else
                        {
                            throw new SystemException("查無預編檔路徑!!");
                        }
                        #endregion

                        #region//銑床預編 MillProgramRequest
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillProgramRequest (MillProgLogId,MillPrePrtData
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillProgRequestId
                                VALUES (@MillProgLogId,@MillPrePrtData
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MillProgLogId,
                                MillPrePrtData = FilePath,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected = insertResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床編程需求紀錄】新增失敗!");

                        #endregion

                        #region//紀錄銑床刀具資訊
                        JObject MillToolDataJson = JObject.Parse(MillToolData);
                        for (int i=0; i< MillToolDataJson["result"].Count();i++) {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.MillProgramTool (MillProgLogId,ToolBlockSeq,ToolBlockName
                                ,MillToolD,MillToolRone,MillToolB,MillToolFL,MillToolL,MillToolHD,MillToolR,MillToolRPM,MillToolFPR
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                 
                                VALUES (@MillProgLogId,@ToolBlockSeq,@ToolBlockName
                                ,@MillToolD,@MillToolRone,@MillToolB,@MillToolFL,@MillToolL,@MillToolHD,@MillToolR,@MillToolRPM,@MillToolFPR
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MillProgLogId,
                                    ToolBlockSeq = MillToolDataJson["result"][i]["ToolBlockSeq"].ToString(),
                                    ToolBlockName = MillToolDataJson["result"][i]["ToolBlockName"].ToString(),
                                    MillToolD = MillToolDataJson["result"][i]["MillToolD"].ToString(),
                                    MillToolRone = MillToolDataJson["result"][i]["MillToolRone"].ToString(),
                                    MillToolB = MillToolDataJson["result"][i]["MillToolB"].ToString(),
                                    MillToolFL = MillToolDataJson["result"][i]["MillToolFL"].ToString(),
                                    MillToolL = MillToolDataJson["result"][i]["MillToolL"].ToString(),
                                    MillToolHD = MillToolDataJson["result"][i]["MillToolHD"].ToString(),
                                    MillToolR = MillToolDataJson["result"][i]["MillToolR"].ToString(),
                                    MillToolRPM = MillToolDataJson["result"][i]["MillToolRPM"].ToString(),
                                    MillToolFPR = MillToolDataJson["result"][i]["MillToolFPR"].ToString(),
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);                           
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = MillProgLogIdResult
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

        #region//AddMillProgramResponsest 新增銑床編程回傳紀錄
        public string AddMillProgramResponsest(int MillProgLogId, string CompanyNo, string ServerPath, string MillPrtData)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //判斷【銑床預編執行紀錄Log】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MillProgLogId,LastModifiedBy
                                 FROM MES.MillProgramLog                            
                                WHERE MillProgLogId =@MillProgLogId";
                        dynamicParameters.Add("MillProgLogId", MillProgLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【銑床預編執行紀錄Log】不存在");
                        int UserId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).LastModifiedBy;
                        #endregion

                        //回傳檔案為多檔的情況
                        string[] ResponsestList = MillPrtData.Split(';');
                        int ResponsestNum = ResponsestList.Count();
                        for (int i = 0; i < ResponsestNum; i++) {
                            #region//檔案處理  
                            string FolderPath = Path.Combine(ServerPath, CompanyNo, "MillProgram", "Responsest");
                            string ASCIIPath = "";
                            string FilePath = "";
                            ASCIIPath = ResponsestList[i].ToString();
                            if (ResponsestList[i].ToString().IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                            if (ResponsestList[i].ToString().IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                            if (ResponsestList[i].ToString().IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                            if (ResponsestList[i].ToString().IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                            if (ResponsestList[i].ToString().IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                            if (ResponsestList[i].ToString().IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");
                            if (ResponsestList[i].ToString().IndexOf("%25") != -1) ASCIIPath = ASCIIPath.Replace("%25", "%");

                            if (File.Exists(ASCIIPath))
                            {
                                string FileName = Path.GetFileNameWithoutExtension(ResponsestList[i].ToString());
                                string FileExtension = Path.GetExtension(ResponsestList[i].ToString());
                                FilePath = Path.Combine(FolderPath, FileName + "(" + MillProgLogId + ")"+ FileExtension);
                                if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                                byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                                File.WriteAllBytes(FilePath, cadFileByte);
                            }
                            else
                            {
                                throw new SystemException("查無預編檔路徑!!");
                            }
                            #endregion

                            #region//銑床預編MillProgramResponsest
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.MillProgramResponsest (MillProgLogId,MillPrtData
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillProgResponsestId
                                VALUES (@MillProgLogId,@MillPrtData
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MillProgLogId,
                                    MillPrtData = FilePath,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            int rowsAffected = insertResult.Count();
                            if (rowsAffected != 1) throw new SystemException("【銑床預編回傳紀錄】新增失敗!");
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

        #region//AddMillProgramWork 新增銑床編程報工紀錄
        public string AddMillProgramWork(int MillProgLogId,int MoProcessId,string MachineNo, string StartWorkDate,string EndWorkDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //判斷【銑床預編執行紀錄Log】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MillProgLogId,LastModifiedBy
                                 FROM MES.MillProgramLog                            
                                WHERE MillProgLogId =@MillProgLogId";
                        dynamicParameters.Add("MillProgLogId", MillProgLogId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【銑床編程執行紀錄Log】不存在");
                        int UserId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).LastModifiedBy;
                        #endregion

                        #region //判斷【製令製程】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MoProcessId
                                 FROM MES.MoProcess                            
                                WHERE MoProcessId =@MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【製令製程】不存在");
                        #endregion

                        #region //判斷【機台】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MachineId
                                 FROM MES.Machine                            
                                WHERE MachineNo =@MachineNo";
                        dynamicParameters.Add("MachineNo", MachineNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【機台】不存在");
                        int MachineId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MachineId;
                        #endregion

                        #region// 銑床編程報工紀錄，只能有一個【銑床編程執行紀錄Log ID】
                        sql = @"SELECT MillProgLogId
                                 FROM MES.MillProgramWork                            
                                WHERE MillProgLogId =@MillProgLogId";
                        dynamicParameters.Add("MillProgLogId", MillProgLogId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 0) throw new SystemException("該【銑床編程報工紀錄，不能有同樣的【銑床編程執行紀錄Log ID】");                     
                        #endregion

                        #region//銑床編程報工紀錄
                        //日期時間相減取秒數
                        DateTime startTime = DateTime.Parse(StartWorkDate);
                        DateTime endTime = DateTime.Parse(EndWorkDate);
                        TimeSpan timeDifference = endTime - startTime;
                        double seconds = timeDifference.TotalSeconds;
                        string WorkCT = seconds.ToString();
                        if (int.Parse(WorkCT) < 0) { throw new SystemException("【銑床編程報工紀錄】"+ WorkCT + "為負數，工時計算失敗!"); }
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MillProgramWork (MillProgLogId,MoProcessId,MachineId,StartWorkDate,EndWorkDate,WorkCT
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MillProgramWorkId
                                VALUES (@MillProgLogId,@MoProcessId,@MachineId,@StartWorkDate,@EndWorkDate,@WorkCT
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MillProgLogId,
                                MoProcessId,
                                MachineId,
                                StartWorkDate,
                                EndWorkDate,
                                WorkCT,
                                Status = "S", //沒有拋轉機台之前的狀態都是S
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();
                        if (rowsAffected != 1) throw new SystemException("【銑床編程報工紀錄】新增失敗!");
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #endregion

        #region//Update

        #region //UpdateCncProgram -- 編程程式基本資料 修改 --Ding 2022.10.04
            public string UpdateCncProgram(int CncProgId, int ProcessId, string CncProgName, string CncProgApi)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region//UpdateCncProgramFile -- 編程程式檔案基本資料 修改 --Ding 2022.10.04
        public string UpdateCncProgramFile(int CncProgramFileId, int CncProgId, string FileDesc, int FileId)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region //UpdateCncParameter -- 編程程式參數基本資料 修改 --Ding 2022.10.04
        public string UpdateCncParameter(int CncParamId, int CncProgId, string CncParamNo, string CncParamName, string CncParamType, string CncParamOption
            , string CncParamUomId, string CncParamValue, string CncParamStatus)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region //UpdateCncProgramKnife -- 編程程式刀具基本資料 修改 --Ding 2022.10.04
        public string UpdateCncProgramKnife(int CncProgramKnifeId, int CncProgId, string KnifeName, string KnifeKey)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();



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

        #region//UpdateCncProgramRequest --編程參數狀態 修改
        public string UpdateCncProgramRequest(int CncProgId,string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.CncProgramRequest SET
                                [Status] = @Status
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE KeyId = @CncProgId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                CncProgId,
                                LastModifiedDate,
                                LastModifiedBy                               
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region//UpdateCncProgramRequest --編程參數狀態 修改
        public string UpdateCncProgramResponsest(int CncProgResponsestId, string ResponsestValues, int FileId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【CncProgResponsestId】的原本的ResponsestValues 
                        string bedoreResponsestValues = "", newResponsestValues="";
                        int BeforeFileId = -1, NewFileId=-1;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ResponsestValues ,a.FileId
                                    FROM MES.CncProgramResponsest a                              
                                    WHERE a.CncProgResponsestId =@CncProgResponsestId";
                        dynamicParameters.Add("CncProgResponsestId", CncProgResponsestId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                            throw new SystemException("該【CncProgResponsestId】不存在");
                        }
                        else {
                            foreach (var item in result)
                            {
                                bedoreResponsestValues = item.ResponsestValues;
                                BeforeFileId = item.FileId;
                            }
                        }
                        if (bedoreResponsestValues.Contains(":")) {
                            string[] txt = bedoreResponsestValues.Split(':');
                            newResponsestValues = ResponsestValues+":"+txt[1].ToString();
                        }
                        #endregion

                        #region 紀錄回傳紀錄  
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.CncResponsestFileLog (CncProgResponsestId,BeforeFileId,NewFileId
                                ,[Status],CreateDate, LastModifiedDate, CreateBy,LastModifiedBy)                                
                                VALUES (@CncProgResponsestId,@BeforeFileId,@NewFileId
                                ,@Status,@CreateDate, @LastModifiedDate, @CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgResponsestId,
                                BeforeFileId,
                                NewFileId= FileId,                              
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.CncProgramResponsest SET
                                ResponsestValues = @newResponsestValues,
                                FileId=@FileId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CncProgResponsestId = @CncProgResponsestId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CncProgResponsestId,
                                newResponsestValues,
                                FileId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region//UpdateToolMachineParameter --機台所需刀具參數 修改
        public string UpdateToolMachineParameter(int ToolMachineId, int ToolClassIdSettingId, string ParameterValue, int MachineId, int MachineLocation, string ToolNo, int AddCount)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【工具】資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToolId
                                FROM MES.Tool a                                
                                WHERE a.ToolNo = @ToolNo";
                        dynamicParameters.Add("ToolNo", ToolNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【工具】資料錯誤!");

                        int ToolId = -1;
                        foreach (var item in result)
                        {
                            ToolId = item.ToolId;
                        }
                        #endregion

                        #region //判斷【機台工具ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolMachine a                                
                                WHERE a.ToolMachineId = @ToolMachineId";
                        dynamicParameters.Add("ToolMachineId", ToolMachineId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【機台工具ID】不存在");
                        #endregion

                        #region //判斷【工具參數ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToolClassIdSettingId 
                                FROM MES.ToolClassIdSetting a                                
                                WHERE a.ToolClassIdSettingId =@ToolClassIdSettingId";
                        dynamicParameters.Add("ToolClassIdSettingId", ToolClassIdSettingId);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("該【工具參數ID】不存在");
                        #endregion

                        int rowsAffected = 0;
                        #region //check
                        if (AddCount == 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.ToolMachine a                              
                                    WHERE a.MachineId = @MachineId
                                    AND a.MachineLocation = @MachineLocation
                                    AND a.ToolMachineId != @ToolMachineId";
                            dynamicParameters.Add("MachineId", MachineId);
                            dynamicParameters.Add("MachineLocation", MachineLocation);
                            dynamicParameters.Add("ToolMachineId", ToolMachineId);
                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() > 0) throw new SystemException("該【刀座】已綁定刀具");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.ToolMachine a                              
                                    WHERE a.ToolId = @ToolId
                                    AND a.ToolMachineId != @ToolMachineId";
                            dynamicParameters.Add("ToolId", ToolId);
                            dynamicParameters.Add("ToolMachineId", ToolMachineId);
                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() > 0) throw new SystemException("該【刀具】已被綁定");

                            #region //UPDATE MES.ToolMachine
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.ToolMachine SET
                                    MachineId = @MachineId,                                
                                    ToolId = @ToolId,                                
                                    MachineLocation = @MachineLocation,                                
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ToolMachineId = @ToolMachineId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MachineId,
                                    ToolId,
                                    MachineLocation,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ToolMachineId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //UPDAET MES.ToolMachineParameter
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ToolMachineParameter SET
                                ParameterValue = @ParameterValue,                                
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ToolMachineId = @ToolMachineId
                                AND ToolClassIdSettingId = @ToolClassIdSettingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ParameterValue,
                                LastModifiedDate,
                                LastModifiedBy,
                                ToolMachineId,
                                ToolClassIdSettingId
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

        #region//UpdateCncProgramWork --報工紀錄 修改
        public string UpdateCncProgramWork(int CncProgramWorkId, string StartWorkDate, string EndWorkDate, string WorkCT)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();


                        #region //判斷報工紀錄是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CncProgramWorkId
                                FROM MES.CncProgramWork  a                                
                                WHERE a.CncProgramWorkId = @CncProgramWorkId";
                        dynamicParameters.Add("CncProgramWorkId", CncProgramWorkId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【編程報工紀錄】資料錯誤!");                       
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.CncProgramWork SET
                                StartWorkDate = @StartWorkDate,
                                EndWorkDate = @EndWorkDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CncProgramWorkId = @CncProgramWorkId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                StartWorkDate,
                                EndWorkDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                CncProgramWorkId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【編程報工紀錄】修改失敗!");
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

        #region//UpdateMillMachineTool --銑床機台刀具 修改
        public string UpdateMillMachineTool(int MillMachineToolId , string ToolNo , int MachineId , string ToolBlockSeq , string ToolBlockName 
            , string MillToolD , string MillToolRone , string MillToolB , string MillToolFL , string MillToolL , string MillToolHD 
            , string MillToolR , string MillToolRPM , string MillToolFPR ,int OriMachineId,string OriToolNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        int OriToolId = -1;

                        #region //判斷【工具ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ToolId
                                FROM MES.Tool                               
                                WHERE ToolNo =@ToolNo";
                        dynamicParameters.Add("ToolNo", ToolNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("該【工具ID】不存在");
                        int ToolId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolId;
                        #endregion

                        
                        if (OriToolNo!="") {
                            #region //判斷【原先 工具ID】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ToolId
                                FROM MES.Tool                               
                                WHERE ToolNo =@OriToolNo";
                            dynamicParameters.Add("OriToolNo", OriToolNo);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() != 1) throw new SystemException("該【原先 工具ID】不存在");
                            OriToolId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolId;
                            #endregion                            
                        }
                        else
                        {
                            OriToolNo = ToolNo;
                        }

                        #region //判斷【機台ID】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine                              
                                WHERE MachineId =@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("該【機台ID】不存在");
                        #endregion

                        #region//同機台不可出現重複的刀座
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM MES.MillMachineTool                              
                                WHERE MachineId =@MachineId
                                AND ToolBlockSeq=@ToolBlockSeq";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("ToolBlockSeq", ToolBlockSeq);
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() != 0 && result3.Count()>1) {
                            throw new SystemException("【該機台不可出現重複的刀座】");
                        }
                        #endregion

                        #region//已經維護的刀具，不可以同時出現在2台機台上
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.MachineDesc,a.MachineId
                                FROM MES.MillMachineTool a
                                INNER JOIN MES.Machine b ON a.MachineId=b.MachineId                            
                                WHERE a.ToolId =@ToolId
                                AND a.MachineId=@MachineId";
                        dynamicParameters.Add("ToolId", ToolId);
                        dynamicParameters.Add("MachineId", MachineId);
                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 1)
                        {                            
                            foreach (var item in result4)
                            {
                                if (MachineId!=item.MachineId) {
                                    throw new SystemException("該刀具目前架設在【" + item.MachineDesc + "】，不可修改");
                                }
                            }                            
                        }
                        #endregion

                        #region//修改銑床機台刀具資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MillMachineTool SET
                                ToolId = @ToolId,
                                MachineId = @MachineId,
                                ToolBlockSeq = @ToolBlockSeq,
                                ToolBlockName = @ToolBlockName,
                                MillToolD = @MillToolD,
                                MillToolRone = @MillToolRone,
                                MillToolB = @MillToolB,
                                MillToolFL = @MillToolFL,
                                MillToolL = @MillToolL,
                                MillToolHD = @MillToolHD,
                                MillToolR = @MillToolR,
                                MillToolRPM = @MillToolRPM,
                                MillToolFPR = @MillToolFPR,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MillMachineToolId = @MillMachineToolId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolId,
                                MachineId,
                                ToolBlockSeq,
                                ToolBlockName,
                                MillToolD,
                                MillToolRone,
                                MillToolB,
                                MillToolFL,
                                MillToolL,
                                MillToolHD,
                                MillToolR,
                                MillToolRPM,
                                MillToolFPR,
                                LastModifiedDate,
                                LastModifiedBy,
                                MillMachineToolId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【銑床機台刀具】修改失敗!");
                        #endregion


                        if (OriMachineId == MachineId && ToolNo != OriToolNo) {
                            string TransactionType = "", ToolLocatorNameNow = "", ToolInventoryNameNow = "";
                            int ToolLocatorIdNow = 0;

                            string TempToolLocatorNo = "TempLocator01";
                            string OriTransactionType = "", OriToolLocatorNameNow = "", OriToolInventoryNameNow = "";
                            int OriToolLocatorIdNow = 0;

                            #region //移除刀具 OricToolNo

                            #region //判斷【工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId,a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorNo = @TempToolLocatorNo";
                            dynamicParameters.Add("TempToolLocatorNo", TempToolLocatorNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                            int TempToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            #endregion

                            #region //判斷工具是否有入庫過
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @OriToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                            dynamicParameters.Add("OriToolId", OriToolId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                if (OriTransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                            }
                            else
                            {
                                OriToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                                OriToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                                OriToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                                if (OriToolLocatorIdNow == TempToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                            }
                            #endregion

                            if (OriToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@OriToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        OriToolId,
                                        TransactionType = "Out",
                                        TransactionDate = DateTime.Now,
                                        ToolLocatorId = OriToolLocatorIdNow,
                                        TraderId = CreateBy,
                                        TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult01.Count();
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@OriToolId, @TransactionType, @TransactionDate, @TempToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OriToolId,
                                    TransactionType = "In",
                                    TransactionDate = DateTime.Now,
                                    TempToolLocatorId,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ProcessingQty = 0
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region//加入刀具 ToolNo
                            #region //判斷【機台工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ToolLocatorId
                                FROM MES.MachineToolLocator                               
                                WHERE MachineId = @MachineId";
                            dynamicParameters.Add("MachineId", MachineId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("機台工具儲位不存在，請重新輸入!");
                            int ToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            #endregion

                            #region //判斷【工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId,a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorId = @ToolLocatorId";
                            dynamicParameters.Add("ToolLocatorId", ToolLocatorId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                            ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            #endregion

                            #region //判斷工具是否有入庫過
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @ToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                            dynamicParameters.Add("ToolId", ToolId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                if (TransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                            }
                            else
                            {
                                ToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                                ToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                                ToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                                if (ToolLocatorIdNow == ToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                            }
                            #endregion

                            if (ToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ToolId,
                                        TransactionType = "Out",
                                        TransactionDate = DateTime.Now,
                                        ToolLocatorId = ToolLocatorIdNow,
                                        TraderId = CreateBy,
                                        TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult01.Count();
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolId,
                                    TransactionType = "In",
                                    TransactionDate = DateTime.Now,
                                    ToolLocatorId,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ProcessingQty = 0
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                        }
                        else if (OriMachineId != MachineId && ToolNo == OriToolNo)
                        {
                            #region //刀具異動
                            string TransactionType = "", ToolLocatorNameNow = "", ToolInventoryNameNow = "";
                            int ToolLocatorIdNow = 0;

                            #region //判斷【機台工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ToolLocatorId
                                FROM MES.MachineToolLocator                               
                                WHERE MachineId = @MachineId";
                            dynamicParameters.Add("MachineId", MachineId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("機台工具儲位不存在，請重新輸入!");
                            int ToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            #endregion

                            #region //判斷【工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorId = @ToolLocatorId";
                            dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                            string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            #endregion

                            #region //判斷工具是否有入庫過
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @ToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                            dynamicParameters.Add("ToolId", ToolId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                if (TransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                            }
                            else
                            {
                                ToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                                ToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                                ToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                                if (ToolLocatorIdNow == ToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                            }
                            #endregion

                            if (ToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ToolId,
                                        TransactionType = "Out",
                                        TransactionDate = DateTime.Now,
                                        ToolLocatorId = ToolLocatorIdNow,
                                        TraderId = CreateBy,
                                        TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult01.Count();
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolId,
                                    TransactionType = "In",
                                    TransactionDate = DateTime.Now,
                                    ToolLocatorId,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ProcessingQty = 0
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                        } else if (OriMachineId != MachineId && ToolNo != OriToolNo) {
                            string TransactionType = "", ToolLocatorNameNow = "", ToolInventoryNameNow = "";
                            int ToolLocatorIdNow = 0;

                            string TempToolLocatorNo = "TempLocator01";
                            string OriTransactionType = "", OriToolLocatorNameNow = "", OriToolInventoryNameNow = "";
                            int OriToolLocatorIdNow = 0;

                            #region //移除刀具 OricToolNo

                            #region //判斷【工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId,a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorNo = @TempToolLocatorNo";
                            dynamicParameters.Add("TempToolLocatorNo", TempToolLocatorNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                            int TempToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            #endregion

                            #region //判斷工具是否有入庫過
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @OriToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                            dynamicParameters.Add("OriToolId", OriToolId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                if (OriTransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                            }
                            else
                            {
                                OriToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                                OriToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                                OriToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                                if (OriToolLocatorIdNow == TempToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                            }
                            #endregion

                            if (OriToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@OriToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        OriToolId,
                                        TransactionType = "Out",
                                        TransactionDate = DateTime.Now,
                                        ToolLocatorId = OriToolLocatorIdNow,
                                        TraderId = CreateBy,
                                        TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult01.Count();
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@OriToolId, @TransactionType, @TransactionDate, @TempToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OriToolId,
                                    TransactionType = "In",
                                    TransactionDate = DateTime.Now,
                                    TempToolLocatorId,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ProcessingQty = 0
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region//加入刀具 ToolNo
                            #region //判斷【機台工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ToolLocatorId
                                FROM MES.MachineToolLocator                               
                                WHERE MachineId = @MachineId";
                            dynamicParameters.Add("MachineId", MachineId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("機台工具儲位不存在，請重新輸入!");
                            int ToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            #endregion

                            #region //判斷【工具儲位】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId,a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorId = @ToolLocatorId";
                            dynamicParameters.Add("ToolLocatorId", ToolLocatorId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                            ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            #endregion

                            #region //判斷工具是否有入庫過
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @ToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                            dynamicParameters.Add("ToolId", ToolId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                if (TransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                            }
                            else
                            {
                                ToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                                ToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                                ToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                                if (ToolLocatorIdNow == ToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                            }
                            #endregion

                            if (ToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ToolId,
                                        TransactionType = "Out",
                                        TransactionDate = DateTime.Now,
                                        ToolLocatorId = ToolLocatorIdNow,
                                        TraderId = CreateBy,
                                        TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult01.Count();
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ToolId,
                                    TransactionType = "In",
                                    TransactionDate = DateTime.Now,
                                    ToolLocatorId,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ProcessingQty = 0
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion
                        } else if (OriMachineId == MachineId && ToolNo == OriToolNo)
                        {
                            //#region //刀具異動
                            //string TransactionType = "", ToolLocatorNameNow = "", ToolInventoryNameNow = "";
                            //int ToolLocatorIdNow = 0;

                            //#region //判斷【機台工具儲位】是否存在
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT ToolLocatorId
                            //    FROM MES.MachineToolLocator                               
                            //    WHERE MachineId = @MachineId";
                            //dynamicParameters.Add("MachineId", MachineId);

                            //result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() <= 0) throw new SystemException("機台工具儲位不存在，請重新輸入!");
                            //int ToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            //#endregion

                            //#region //判斷【工具儲位】是否存在
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT TOP 1 a.ToolLocatorName,b.ToolInventoryName
                            //    FROM MES.ToolLocator a
                            //    INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                            //    WHERE ToolLocatorId = @ToolLocatorId";
                            //dynamicParameters.Add("ToolLocatorId", ToolLocatorId);

                            //result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                            //string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            //string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            //#endregion

                            //#region //判斷工具是否有入庫過
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                            //    FROM MES.ToolTransactions a
                            //    INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                            //    INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                            //    WHERE a.ToolId = @ToolId
                            //    AND a.TransactionType = 'In'
                            //    Order By a.CreateDate DESC";
                            //dynamicParameters.Add("ToolId", ToolId);

                            //result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() <= 0)
                            //{
                            //    if (TransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                            //}
                            //else
                            //{
                            //    ToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            //    ToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            //    ToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            //    if (ToolLocatorIdNow == ToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                            //}
                            //#endregion

                            //if (ToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                            //{
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                            //    , TraderId, TransactionReason
                            //    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                            //    OUTPUT INSERTED.ToolTransactionsId
                            //    VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                            //    , @TraderId, @TransactionReason
                            //    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            //    dynamicParameters.AddDynamicParams(
                            //        new
                            //        {
                            //            ToolId,
                            //            TransactionType = "Out",
                            //            TransactionDate = DateTime.Now,
                            //            ToolLocatorId = ToolLocatorIdNow,
                            //            TraderId = CreateBy,
                            //            TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                            //            CreateDate,
                            //            LastModifiedDate,
                            //            CreateBy,
                            //            LastModifiedBy
                            //        });
                            //    var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                            //    rowsAffected += insertResult01.Count();
                            //}

                            //dynamicParameters = new DynamicParameters();
                            //sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                            //    , TraderId, TransactionReason
                            //    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                            //    OUTPUT INSERTED.ToolTransactionsId
                            //    VALUES (@ToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                            //    , @TraderId, @TransactionReason
                            //    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                            //dynamicParameters.AddDynamicParams(
                            //    new
                            //    {
                            //        ToolId,
                            //        TransactionType = "In",
                            //        TransactionDate = DateTime.Now,
                            //        ToolLocatorId,
                            //        TraderId = CreateBy,
                            //        TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                            //        CreateDate,
                            //        LastModifiedDate,
                            //        CreateBy,
                            //        LastModifiedBy,
                            //        ProcessingQty = 0
                            //    });
                            //var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            //#endregion
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
            }
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

        #region//Delete
        #region//DeleteToolMachineParameter --機台所需刀具參數 刪除
        public string DeleteToolMachineParameter(int ToolMachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【工具機台資料】資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ToolMachine a                                
                                WHERE a.ToolMachineId = @ToolMachineId";
                        dynamicParameters.Add("ToolMachineId", ToolMachineId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【工具機台】資料不存在!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯TABLE                      
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.ToolMachineParameter
                                WHERE ToolMachineId = @ToolMachineId";
                        dynamicParameters.Add("ToolMachineId", ToolMachineId);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除                        
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ToolMachine
                                WHERE ToolMachineId = @ToolMachineId";
                        dynamicParameters.Add("ToolMachineId", ToolMachineId);
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

        #region//DeleteCncProgramWork --編程報工紀錄 刪除
        public string DeleteCncProgramWork(int CncProgramWorkId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷【編程報工紀錄】資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.CncProgramWork a                                
                                WHERE a.CncProgramWorkId = @CncProgramWorkId";
                        dynamicParameters.Add("CncProgramWorkId", CncProgramWorkId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該【編程報工紀錄】資料不存在!");
                        #endregion

                        #region //刪除 
                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.CncProgramWorkFiles
                                WHERE CncProgramWorkId = @CncProgramWorkId";
                        dynamicParameters.Add("CncProgramWorkId", CncProgramWorkId);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected <= 0) throw new SystemException("該【編程報工檔案上傳紀錄】刪除失敗!");
                        #endregion

                        #region //刪除  
                        rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.CncProgramWork
                                WHERE CncProgramWorkId = @CncProgramWorkId";
                        dynamicParameters.Add("CncProgramWorkId", CncProgramWorkId);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected <= 0) throw new SystemException("該【編程報工紀錄】刪除失敗!");
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

        #endregion]

        #region//DeleteMillMachineTool  --銑床刀具刪除
        public string DeleteMillMachineTool(int MillMachineToolId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int? VirtualToolId = null;

                        #region //取出原先的刀具資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ToolId,MachineId
                                FROM MES.MillMachineTool                             
                                WHERE MillMachineToolId =@MillMachineToolId";
                        dynamicParameters.Add("MillMachineToolId", MillMachineToolId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        if (result1.Count() != 1) throw new SystemException("該【工具ID】不存在");
                        int OriToolId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolId;
                        int MachineId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MachineId;
                        #endregion

                        #region//修改銑床機台刀具資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MillMachineTool SET
                                ToolId = @VirtualToolId,                                
                                ToolBlockSeq = @ToolBlockSeq,
                                ToolBlockName = @ToolBlockName,
                                MillToolD = @MillToolD,
                                MillToolRone = @MillToolRone,
                                MillToolB = @MillToolB,
                                MillToolFL = @MillToolFL,
                                MillToolL = @MillToolL,
                                MillToolHD = @MillToolHD,
                                MillToolR = @MillToolR,
                                MillToolRPM = @MillToolRPM,
                                MillToolFPR = @MillToolFPR,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MillMachineToolId = @MillMachineToolId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                VirtualToolId,
                                ToolBlockSeq= "-1",
                                ToolBlockName = "-1",
                                MillToolD = "-1",
                                MillToolRone = "-1",
                                MillToolB = "-1",
                                MillToolFL = "-1",
                                MillToolL = "-1",
                                MillToolHD = "-1",
                                MillToolR = "-1",
                                MillToolRPM = "-1",
                                MillToolFPR = "-1",
                                LastModifiedDate,
                                LastModifiedBy,
                                MillMachineToolId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【銑床機台刀具】移除失敗!");
                        #endregion

                        #region //刀具異動
                        string TransactionType = "", ToolLocatorNameNow = "", ToolInventoryNameNow = "";
                        int ToolLocatorIdNow = 0;

                        string TempToolLocatorNo = "TempLocator01";
                        string OriTransactionType = "", OriToolLocatorNameNow = "", OriToolInventoryNameNow = "";
                        int OriToolLocatorIdNow = 0;

                        #region //移除刀具 OricToolNo

                        #region //判斷【工具儲位】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ToolLocatorId,a.ToolLocatorName,b.ToolInventoryName
                                FROM MES.ToolLocator a
                                INNER JOIN MES.ToolInventory b on a.ToolInventoryId = b.ToolInventoryId
                                WHERE ToolLocatorNo = @TempToolLocatorNo";
                        dynamicParameters.Add("TempToolLocatorNo", TempToolLocatorNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("工具儲位不存在，請重新輸入!");
                        int TempToolLocatorId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                        string ToolLocatorName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                        string ToolInventoryName = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                        #endregion

                        #region //判斷工具是否有入庫過
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ToolLocatorId ,b.ToolLocatorName,c.ToolInventoryName
                                FROM MES.ToolTransactions a
                                INNER JOIN MES.ToolLocator b on a.ToolLocatorId = b.ToolLocatorId
                                INNER JOIN MES.ToolInventory c on b.ToolInventoryId = c.ToolInventoryId
                                WHERE a.ToolId = @OriToolId
                                AND a.TransactionType = 'In'
                                Order By a.CreateDate DESC";
                        dynamicParameters.Add("OriToolId", OriToolId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                            if (OriTransactionType == "Out") throw new SystemException("該工具未曾入庫過,請先做【入庫】才能做【出庫】");
                        }
                        else
                        {
                            OriToolLocatorIdNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorId;
                            OriToolLocatorNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolLocatorName;
                            OriToolInventoryNameNow = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).ToolInventoryName;
                            if (OriToolLocatorIdNow == TempToolLocatorId) throw new SystemException("目前儲位和欲入庫的儲位不可以相同");
                        }
                        #endregion

                        if (OriToolLocatorIdNow > 0) //有現在刀具儲位位置新增出庫紀錄
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@OriToolId, @TransactionType, @TransactionDate, @ToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OriToolId,
                                    TransactionType = "Out",
                                    TransactionDate = DateTime.Now,
                                    ToolLocatorId = OriToolLocatorIdNow,
                                    TraderId = CreateBy,
                                    TransactionReason = CreateDate + "從 【庫房 -" + ToolInventoryNameNow + "】的【儲位 -" + ToolLocatorNameNow + "】移出工具",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult01 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult01.Count();
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ToolTransactions (ToolId, TransactionType, TransactionDate, ToolLocatorId
                                , TraderId, TransactionReason
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy,ProcessingQty)
                                OUTPUT INSERTED.ToolTransactionsId
                                VALUES (@OriToolId, @TransactionType, @TransactionDate, @TempToolLocatorId
                                , @TraderId, @TransactionReason
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ProcessingQty)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OriToolId,
                                TransactionType = "In",
                                TransactionDate = DateTime.Now,
                                TempToolLocatorId,
                                TraderId = CreateBy,
                                TransactionReason = CreateDate + " - 將工具入庫到 【庫房 -" + ToolInventoryName + "】的【儲位 -" + ToolLocatorName + "】",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy,
                                ProcessingQty = 0
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

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
        #endregion



    }
}
