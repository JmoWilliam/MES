using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Text.RegularExpressions;
using System.Text;
using System.Collections.Generic;
using System.Web.UI.WebControls; 
using System.Globalization;
using System.Threading.Tasks;
using System.Data;

namespace MESDA
{
    public class MesBasicInformationDA
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


        public MesBasicInformationDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
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

        #region //Get
        #region //GetCompanyProdMode -- 取得特定公司的生產模式 -- Ding 2022.10.06
        public string GetCompanyProdMode(int CompanyId, int ModeId, string ModeNo, string ModeName, string Status, string BarcodeCtrl
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ModeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.CompanyId, a.ModeNo, a.ModeName, a.ModeDesc, a.Status ,a.BarcodeCtrl
                            , (a.ModeNo + '-' + a.ModeName) txtProdModeWithText
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProdMode a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeNo", @" AND a.ModeNo LIKE '%' + @ModeNo + '%'", ModeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeName", @" AND (a.ModeName LIKE '%' + @ModeName + '%' OR a.ModeDesc LIKE '%' + @ModeName + '%')", ModeName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (BarcodeCtrl.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeCtrl", @" AND a.BarcodeCtrl IN @BarcodeCtrl", BarcodeCtrl.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ModeId DESC";
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

        #region //GetProdMode -- 取得生產模式 -- Zony 2022.05.30
        public string GetProdMode(int ModeId, string ModeNo, string ModeName, string Status, string BarcodeCtrl, string ScrapRegister
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ModeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.CompanyId, a.ModeNo, a.ModeName, a.ModeDesc, a.Status ,a.BarcodeCtrl, a.ScrapRegister
                           ,a.Source ,a.PVTQCFlag ,a.NgToBarcode ,a.TrayBarcode ,a.LotStatus ,a.OutputBarcodeFlag ,a.MrType ,a.OQcCheckType
                           ,(a.ModeNo + '-' + a.ModeName) txtProdModeWithText
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProdMode a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeNo", @" AND a.ModeNo LIKE '%' + @ModeNo + '%'", ModeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeName", @" AND (a.ModeName LIKE '%' + @ModeName + '%' OR a.ModeDesc LIKE '%' + @ModeName + '%')", ModeName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (BarcodeCtrl.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeCtrl", @" AND a.BarcodeCtrl IN @BarcodeCtrl", BarcodeCtrl.Split(','));
                    if (ScrapRegister.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ScrapRegister", @" AND a.ScrapRegister IN @ScrapRegister", ScrapRegister.Split(','));
                    

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ModeId";
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

        #region //GetProdModeShift -- 取得生產模式班次資料 -- Shintokuro 2022.06.22
        public string GetProdModeShift(int ModeShiftId, int ModeId, int ShiftId, string EffectiveStartDate, string EffectiveEndDate
            , string ExpirationStartDate, string ExpirationEndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ModeShiftId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , FORMAT(a.EffectiveDate, 'yyyy-MM-dd') EffectiveDate, FORMAT(a.ExpirationDate, 'yyyy-MM-dd') ExpirationDate, a.Status
                            , b.ModeId, b.ModeNo, b.ModeName
                            , c.ShiftId, c.ShiftName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProdModeShift a
                          INNER JOIN MES.ProdMode b on a.ModeId = b.ModeId
                          INNER JOIN BAS.[Shift] c on a.ShiftId = c.ShiftId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeShiftId", @" AND a.ModeShiftId = @ModeShiftId", ModeShiftId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShiftId", @" AND a.ShiftId = @ShiftId", ShiftId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EffectiveStartDate", @" AND a.EffectiveDate >= @EffectiveStartDate ", EffectiveStartDate.Length > 0 ? Convert.ToDateTime(EffectiveStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EffectiveEndDate", @" AND a.EffectiveDate <= @EffectiveEndDate ", EffectiveEndDate.Length > 0 ? Convert.ToDateTime(EffectiveEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ExpirationStartDate", @" AND a.ExpirationDate >= @ExpirationStartDate ", ExpirationStartDate.Length > 0 ? Convert.ToDateTime(ExpirationStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ExpirationEndDate", @" AND a.ExpirationDate <= @ExpirationEndDate ", ExpirationEndDate.Length > 0 ? Convert.ToDateTime(ExpirationEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));



                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ModeShiftId DESC";
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

        #region //GetWorkShop -- 取得車間資料 -- Shintokuro 2022.06.07
        public string GetWorkShop(int ShopId, string ShopNo, string ShopName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ShopId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.ShopNo, a.ShopName, a.ShopDesc, a.Location, a.Floor, a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.WorkShop a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopNo", @" AND a.ShopNo LIKE '%' + @ShopNo + '%'", ShopNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopName", @" AND a.ShopName LIKE '%' + @ShopName + '%'", ShopName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ShopId DESC";
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

        #region //GetMachine -- 取得機台資料 -- Shintokuro 2022.06.07
        public string GetMachine(int MachineId, int ShopId, string MachineNo, string MachineName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MachineId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.MachineNo, a.MachineName, a.MachineDesc, a.Status,  (a.MachineName +'(' + a.MachineDesc +'/'+b.ShopName+')')  as MachineDescFull
                           ,  (a.MachineName +'(' + a.MachineDesc +')')  as MachineNameDesc
                           , b.ShopId, b.ShopName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.Machine a
                          LEFT JOIN MES.WorkShop b on b.ShopId = a.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineNo", @" AND a.MachineNo LIKE '%' + @MachineNo + '%'", MachineNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineName", @" AND a.MachineName LIKE '%' + @MachineName + '%'", MachineName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MachineId DESC";
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

        #region //GetWorkingMachine -- 取得加工機台資料 -- Ding 2022.12.03
        public string GetWorkingMachine(int MachineId, int ShopId, string MachineNo, string MachineName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MachineId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.MachineNo, a.MachineName, a.MachineDesc, a.Status
                           , b.ShopId, b.ShopName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.Machine a
                          LEFT JOIN MES.WorkShop b on b.ShopId = a.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"
                            AND NOT exists(select 1
                                from QMS.QmmDetail b
                                where a.MachineId = b.MachineId)
                            AND b.CompanyId = @CompanyId
                            ";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineNo", @" AND a.MachineNo LIKE '%' + @MachineNo + '%'", MachineNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineName", @" AND a.MachineName LIKE '%' + @MachineName + '%'", MachineName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MachineId DESC";
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

        #region //GetMachineAsset -- 取得機台資產編號 -- Shintokuro 2022.06.10
        public string GetMachineAsset(int MachineId, string AssetNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MachineId, a.AssetNo";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.MachineId, a.AssetNo
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MachineAsset a
                          INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                          INNER JOIN MES.WorkShop c on b.ShopId = b.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AssetNo", @" AND a.AssetNo LIKE '%' + @AssetNo + '%'", AssetNo);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MachineId DESC";
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

        #region //GetMachineLog -- 取得機台日誌 -- Shintokuro 2022.06.13
        public string GetMachineLog(int MachineId, string StartDate, string EndDate, string OperatingStatus, int UserId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.LogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.BeginTime, a.EndTime, a.OperationTime, a.OperatingStatus
                            , b.MachineId, b.MachineName
                            , d.UserId, d.UserName
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MachineLog a
                          INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                          INNER JOIN MES.WorkShop c ON b.ShopId = c.ShopId
                          INNER JOIN BAS.[User] d ON a.Operator = d.UserId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OperatingStatus", @" AND a.OperatingStatus = @OperatingStatus", OperatingStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Operator", @" AND a.Operator = @Operator", UserId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LogId DESC";
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

        #region //GetTimeInterval -- 取得時間區段資料 -- Shintokuro 2022.06.15
        public string GetTimeInterval(int IntervalId, string BeginTime, string EndTime
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.IntervalId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , CONVERT(VARCHAR(5), a.BeginTime, 108) BeginTime, CONVERT(VARCHAR(5), a.EndTime, 108) EndTime
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.TimeInterval a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "IntervalId", @" AND a.IntervalId = @IntervalId", IntervalId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BeginTime", @" AND a.BeginTime >= @BeginTime", BeginTime);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndTime", @" AND a.EndTime <= @EndTime", EndTime);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.IntervalId DESC";
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

        #region //GetDevice -- 取得裝置資料 -- Shintokuro 2022.06.15
        public string GetDevice(int DeviceId, string DeviceType, string DeviceName, string DeviceIdentifierCode, string DeviceAuthority, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.DeviceId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.DeviceName, a.DeviceType, a.DeviceName, a.DeviceDesc, a.DeviceIdentifierCode, a.DeviceAuthority, a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.Device a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceId", @" AND a.DeviceId = @DeviceId", DeviceId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceType", @" AND a.DeviceType = @DeviceType", DeviceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceName", @" AND a.DeviceName LIKE '%' + @DeviceName + '%'" , DeviceName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceIdentifierCode", @" AND a.DeviceIdentifierCode LIKE '%' + @DeviceIdentifierCode + '%'", DeviceIdentifierCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceAuthority", @" AND a.DeviceAuthority = @DeviceAuthority", DeviceAuthority);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DeviceId DESC";
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

        #region //GetDeviceMachine -- 取得裝置機台綁定資料 -- Shintokuro 2022.06.16
        public string GetDeviceMachine(int DeviceId, int MachineId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.DeviceId,a.MachineId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", b.MachineName, b.MachineDesc";
                    sqlQuery.mainTables =
                        @"FROM MES.DeviceMachine a
                          LEFT JOIN MES.Machine b on a.MachineId =b.MachineId
                          INNER JOIN MES.Device c on a.DeviceId =c.DeviceId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceId", @" AND a.DeviceId = @DeviceId", DeviceId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DeviceId DESC";
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

        #region //GetProcess -- 取得製程資料 -- Shintokuro 2022.06.20
        public string GetProcess(int ProcessId, string ProcessNo, string ProcessName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProcessNo, a.ProcessName, a.ProcessDesc, a.Status, (a.ProcessNo + '-' + a.ProcessName) ProcessWithText
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.Process a 
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND a.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessNo", @" AND a.ProcessNo LIKE '%' + @ProcessNo + '%'", ProcessNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessName", @" AND a.ProcessName LIKE '%' + @ProcessName + '%'", ProcessName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProcessId";
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

        #region //GetProdUnit -- 取得生產單元資料 -- Shintokuro 2022.06.23
        public string GetProdUnit(int UnitId, string UnitNo, string UnitName, string CheckStatus, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.UnitId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UnitNo, a.UnitName, a.UnitDesc, a.CheckStatus, a.Status
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProdUnit a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UnitId", @" AND a.UnitId = @UnitId", UnitId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UnitNo", @" AND a.UnitNo LIKE '%' + @UnitNo + '%'", UnitNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UnitName", @" AND a.UnitName LIKE '%' + @UnitName + '%'", UnitName);
                    if (CheckStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CheckStatus", @" AND a.CheckStatus IN @CheckStatus", CheckStatus.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UnitId DESC";
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

        #region //GetMachine -- 取得機台資訊資料(過站平板用) -- William 2022-06-29
        public string GetMachineByDevice(int ShopId, string DeviceIdentifierCode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MachineId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MachineNo, a.MachineName, a.MachineDesc
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.Machine a
                          INNER JOIN MES.WorkShop b on a.ShopId = b.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    if (DeviceIdentifierCode != "")
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeviceIdentifierCode",
                            @"  AND EXISTS(SELECT 1
                                             FROM MES.DeviceMachine b,MES.Device c
                                            WHERE b.DeviceId = c.DeviceId
                                              AND a.MachineId = b.MachineId
                                              AND c.DeviceIdentifierCode = @DeviceIdentifierCode) ",
                            DeviceIdentifierCode);
                    }

                    sqlQuery.conditions = queryCondition;

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

        #region //GetWarehouse -- 取得庫房基本資料 -- Shintokuro 2022.07.01
        public string GetWarehouse(int WarehouseId, int ShopId, string WarehouseNo, string WarehouseName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.WarehouseId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",  a.ShopId, a.WarehouseNo, a.WarehouseName, a.WarehouseDesc, a.Status
                          , b.ShopName
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.Warehouse a
                          INNER JOIN MES.WorkShop b on a.ShopId = b.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WarehouseId", @" AND a.WarehouseId = @WarehouseId", WarehouseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WarehouseNo", @" AND a.WarehouseNo LIKE '%' + @WarehouseNo + '%'", WarehouseNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WarehouseName", @" AND a.WarehouseName LIKE '%' + @WarehouseName + '%'", WarehouseName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.WarehouseId DESC";
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

        #region //GetWarehouseLocation -- 取得庫房儲位基本資料 -- Shintokuro 2022.07.04
        public string GetWarehouseLocation(int WarehouseId, int LocationId, string LocationNo, string LocationName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.LocationId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.WarehouseId, a.LocationNo, a.LocationName, a.LocationDesc, a.Status
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.WarehouseLocation a
                          INNER JOIN MES.Warehouse b on a.WarehouseId = b.WarehouseId
                          INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WarehouseId", @" AND a.WarehouseId = @WarehouseId", WarehouseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LocationId", @" AND a.LocationId = @LocationId", LocationId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LocationNo", @" AND a.LocationNo LIKE '%' + @LocationNo + '%'", LocationNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LocationName", @" AND a.LocationName LIKE '%' + @LocationName + '%'", LocationName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LocationId DESC";
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

        #region //GetProcessParameter -- 取得製程參數資料 -- Shintokuro 2022.07.05
        public string GetProcessParameter(int ParameterId, int ProcessId, int DepartmentId, int ModeId, string ProcessCheckStatus, string PreCollectionStatus
            , string PostCollectionStatus, string NgToBarcode, string PassingMode, string ProcessCheckType
            , string Status, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ParameterId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DepartmentId, ISNULL(d.DepartmentName,'') as DepartmentName, a.ProcessCheckStatus, a.PreCollectionStatus, a.PostCollectionStatus, a.NgToBarcode, a.PassingMode, a.ProcessCheckType, a.Status, a.PackageFlag, a.ConsumeFlag
                          , b.ProcessId, b.ProcessName
                          , c.ModeId, c.ModeName
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProcessParameter a
                            LEFT JOIN MES.Process b on a.ProcessId = b.ProcessId
                            LEFT JOIN MES.ProdMode c on a.ModeId = c.ModeId
                            LEFT JOIN BAS.Department d on d.DepartmentId = a.DepartmentId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterId", @" AND a.ParameterId = @ParameterId", ParameterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND b.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND c.ModeId = @ModeId", ModeId);
                    if (ProcessCheckStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessCheckStatus", @" AND a.ProcessCheckStatus IN @ProcessCheckStatus", ProcessCheckStatus.Split(','));
                    if (PreCollectionStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PreCollectionStatus", @" AND a.PreCollectionStatus IN @PreCollectionStatus", PreCollectionStatus.Split(','));
                    if (PostCollectionStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PostCollectionStatus", @" AND a.PostCollectionStatus IN @PostCollectionStatus", PostCollectionStatus.Split(','));
                    if (NgToBarcode.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "NgToBarcode", @" AND a.NgToBarcode IN @NgToBarcode", NgToBarcode.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PassingMode", @" AND a.PassingMode = @PassingMode", PassingMode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessCheckType", @" AND a.ProcessCheckType = @ProcessCheckType", ProcessCheckType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ParameterId DESC";
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

        #region //GetProcessMachine -- 取得製程機台資料 -- Shintokuro 2022.07.06
        public string GetProcessMachine(int ParameterId, int MachineId, string BatchStatus, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ParameterId, a.MachineId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.BatchStatus, a.ToolCount, a.KeyenceFlag, a.KeyenceId, a.[Status]
                          , b.MachineDesc,b.MachineName
                          , c.ShopName
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProcessMachine a
                          INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                          INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterId", @" AND a.ParameterId = @ParameterId", ParameterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
                    if (BatchStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BatchStatus", @" AND a.BatchStatus IN @BatchStatus", BatchStatus.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ParameterId";
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

        #region //GetProcessProductionUnit -- 取得製程生產單元資料 -- Shintokuro 2022.07.06
        public string GetProcessProductionUnit(int ParameterId, int UnitId, int SortNumber, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ParameterId, a.UnitId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SortNumber, a.Status
                          , b.UnitId, b.UnitName
                         ";
                    sqlQuery.mainTables =
                        @"FROM MES.ProcessProductionUnit a
                          INNER JOIN MES.ProdUnit b on a.UnitId = b.UnitId
                          INNER JOIN MES.ProcessParameter c on a.ParameterId = c.ParameterId
                          INNER JOIN MES.Process d on c.ProcessId = d.ProcessId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterId", @" AND a.ParameterId = @ParameterId", ParameterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UnitId", @" AND a.UnitId = @UnitId", UnitId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SortNumber", @" AND a.SortNumber = @SortNumber", SortNumber);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ParameterId DESC";
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

        #region //GetTray -- 取得托盤資料 -- Shintokuro 2023.02.04
        public string GetTray(int TrayId, string MoNo, string BarcodeNo, string TrayNo, string TrayName, string TrayType, string BindStatus, string Status, string DeleteBatch
            , string CreateUserId, string CreatTimeStart, string CreatTimeEnd
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.TrayId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.TrayNo, a.TrayName, a.TrayCapacity, a.Remark, a.UseTimes, a.Status
                           , a.BarcodeNo
                           ,(e.WoErpPrefix + '-' + e.WoErpNo + '(' + CONVERT ( Varchar , d.WoSeq ) + ')') MoNo
                           ,b.BarcodeQty
                           ,b1.TrayBarcodeLogId,a1.UserNo +':'+a1.UserName CreateUser,FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.Tray a
                          INNER JOIN BAS.[User] a1 on a.CreateBy = a1.UserId
                          LEFT JOIN MES.BarcodePrint b on a.BarcodeNo = b.BarcodeNo
                          LEFT JOIN MES.MoSetting c on b.MoSettingId = c.MoSettingId
                          LEFT JOIN MES.ManufactureOrder d on c.MoId = d.MoId
                          LEFT JOIN MES.WipOrder e on d.WoId = e.WoId
                          OUTER APPLY (SELECT TOP 1 b.TrayBarcodeLogId FROM MES.TrayBarcodeLog b WHERE a.TrayId = b.TrayId) b1
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayId", @" AND a.TrayId = @TrayId", TrayId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoNo", @" AND (e.WoErpPrefix + '-' +  e.WoErpNo + '('+ CONVERT ( Varchar , d.WoSeq ) +')') LIKE '%' + @MoNo +'%'", MoNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeNo", @" AND a.BarcodeNo LIKE '%' + @BarcodeNo + '%'", BarcodeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayNo", @" AND a.TrayNo LIKE '%' + @TrayNo + '%'", TrayNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayName", @" AND a.TrayName LIKE '%' + @TrayName + '%'", TrayName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayType", @" AND a.TrayType = @TrayType", TrayType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CreateUserId", @" AND a.CreateBy = @CreateUserId", CreateUserId);
                    if (BindStatus =="Y") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BindStatus", @" AND a.BarcodeNo is not null", BindStatus);
                    if (BindStatus =="N") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BindStatus", @" AND a.BarcodeNo is null", BindStatus);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (DeleteBatch == "Y") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeleteBatch", @" AND b1.TrayBarcodeLogId is null", DeleteBatch);
                    if (!string.IsNullOrEmpty(CreatTimeStart) && !string.IsNullOrEmpty(CreatTimeEnd))
                    {
                        DateTime st = DateTime.ParseExact(CreatTimeStart, "yyyy-MM-ddTHH:mm", CultureInfo.InvariantCulture);
                        DateTime et = DateTime.ParseExact(CreatTimeEnd, "yyyy-MM-ddTHH:mm", CultureInfo.InvariantCulture)
                                             .AddMinutes(1).AddMilliseconds(-1); // 取該分鐘的最後一毫秒

                        string stStr = st.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string etStr = et.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        queryCondition += $" AND a.CreateDate BETWEEN '{stStr}' AND '{etStr}'";

                        dynamicParameters.Add("CreatTimeStart", st);
                        dynamicParameters.Add("CreatTimeEnd", et);
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TrayId DESC";
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

        #region //GetTrayBarcodeLog -- 取得托盤曾經綁定條碼資料 -- Shintokuro 2023.02.04
        public string GetTrayBarcodeLog(int TrayId, string TrayNo, string TrayNoList
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.TrayBarcodeLogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.TrayId, a.BarcodeNo, FORMAT(a.BindDate, 'yyyy-MM-dd') BindDate, FORMAT(a.ReMoveBindDate, 'yyyy-MM-dd') ReMoveBindDate
                          , b.BarcodeQty
                          ,(e.WoErpPrefix + '-' + e.WoErpNo + '(' + CONVERT ( Varchar , d.WoSeq ) + ')') MoNo
                          ,f.TrayNo";
                    sqlQuery.mainTables =
                        @"FROM MES.TrayBarcodeLog a
                          INNER JOIN MES.BarcodePrint b on a.BarcodeNo = b.BarcodeNo
                          INNER JOIN MES.MoSetting c on b.MoSettingId = c.MoSettingId
                          INNER JOIN MES.ManufactureOrder d on c.MoId = d.MoId
                          INNER JOIN MES.WipOrder e on d.WoId = e.WoId
                          INNER JOIN MES.Tray f on a.TrayId = f.TrayId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND f.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayId", @" AND a.TrayId = @TrayId", TrayId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayNo", @" AND f.TrayNo = @TrayNo", TrayNo);
                    if(TrayNoList != "") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TrayNoList", @" AND f.TrayNo in @TrayNoList", TrayNoList.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TrayBarcodeLogId DESC";
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

        #region //GetKeyence -- 取得Keyence資料 -- Ann 2023-05-26
        public string GetKeyence(int KeyenceId, string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.KeyenceId, a.KeyenceNo, a.KeyenceDesc, a.KeyenceIP, a.KeyencePort, a.[Status]
                            FROM MES.Keyence a
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "KeyenceId", @" AND a.KeyenceId = @KeyenceId", KeyenceId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);

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

        #region //GetUserEventSetting -- 取得人員事件 -- Xuan 2023.07.28
        public string GetUserEventSetting(int UserEventItemId, string UserEventItemNo, string UserEventItemName, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.UserEventItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.UserEventItemNo, a.UserEventItemName, a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.UserEventItem a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserEventItemId", @" AND a.UserEventItemId = @UserEventItemId", UserEventItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserEventItemNo", @" AND a.UserEventItemNo LIKE '%' + @UserEventItemNo + '%'", UserEventItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserEventItemName", @" AND a.UserEventItemName LIKE '%' + REPLACE(@UserEventItemName, ' ', '%') + '%'", UserEventItemName);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserEventItemId DESC";
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

        #region //GetMachineEventSetting -- 取得機台事件 -- Xuan 2023.08.11
        public string GetMachineEventSetting(int ShopId, int MachineId, string ShopName, string MachineName, int MachineEventItemId, string MachineEventNo, string MachineEventName, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.MachineEventItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.MachineEventNo, a.MachineEventName, a.Status, a.MachineId
                          , b.MachineName, b.MachineDesc
                          , c.ShopId, c.ShopName";
                    sqlQuery.mainTables =
                        @"FROM MES.MachineEventItem a
						  INNER JOIN MES.Machine b on b.MachineId=a.MachineId
						  INNER JOIN MES.WorkShop c on c.ShopId=b.ShopId";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineEventItemId", @" AND a.MachineEventItemId = @MachineEventItemId", MachineEventItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineEventNo", @" AND a.MachineEventNo LIKE '%' + @MachineEventNo + '%'", MachineEventNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND c.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND b.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopName", @" AND c.ShopName LIKE '%' + @ShopName + '%'", ShopName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineName", @" AND b.MachineName LIKE '%' + @MachineName + '%'", MachineName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineEventName", @" AND a.MachineEventName LIKE '%' + @MachineEventName + '%'", MachineEventName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MachineEventItemId DESC";
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

        #region //GetProcessEventSetting -- 取得加工事件 -- Xuan 2023.08.15
        public string GetProcessEventSetting(int ProcessEventItemId, int ProcessId, string TypeNo, string ProcessEventName, int ParameterId, int ModeId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ProcessEventItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",c.CompanyId
                          ,a.ParameterId
                          ,b.ModeId
                          ,e.ModeName
                          ,c.ProcessName
                          ,a.ProcessEventType
                          ,d.TypeNo
                          ,d.TypeName
                          ,a.ProcessEventName
                          ,a.Status
                          ,a.CreateDate
                          ,a.LastModifiedDate
                          ,a.CreateBy
                          ,a.LastModifiedBy";
                    sqlQuery.mainTables =
                        @"FROM MES.ProcessEventItem a
                          INNER JOIN MES.ProcessParameter b on b.ParameterId=a.ParameterId
                          INNER JOIN MES.Process c on c.ProcessId=b.ProcessId
                          INNER JOIN BAS.Type d on d.TypeNo=a.ProcessEventType and d.TypeSchema='ProcessEventItem'
                          INNER JOIN MES.ProdMode e on e.ModeId=b.ModeId";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessEventItemId", @" AND a.ProcessEventItemId = @ProcessEventItemId", ProcessEventItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND c.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND b.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterId", @" AND b.ParameterId = @ParameterId", ParameterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeNo", @" AND d.TypeNo = @TypeNo", TypeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessEventName", @" AND a.ProcessEventName LIKE '%' + @ProcessEventName + '%'", ProcessEventName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProcessEventItemId DESC";
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

        #region //GetProcessParameterQcItem //取得製程參數量測項目 --GPai 2024 0111
        public string GetProcessParameterQcItem(int ParameterId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.ParameterQcItemId, a.ParameterId, a.QcItemId, b.QcItemNo, b.QcItemName, a.QcItemDesc, a.DesignValue, a.UpperTolerance, a.LowerTolerance, a.Remark
                            , a.BallMark, a.QmmDetailId, a.Unit, d.MachineName, c.MachineNumber, d.MachineDesc
                            FROM MES.ProcessParameterQcItem a
                            INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
							LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
							LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                            WHERE a.ParameterId = @ParameterId";
                    dynamicParameters.Add("ParameterId", ParameterId);

                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);

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

        #region //GetRoutingItemQcItem //途程量測項目 --GPai 2024 0111
        public string GetRoutingItemQcItem(int RoutingItemId, int ItemProcessId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.RoutingItemQcItemId, a.RoutingItemId, a.QcItemId, b.QcItemNo, b.QcItemName, a.QcItemDesc, a.DesignValue, a.UpperTolerance, a.LowerTolerance, a.Remark, a.ItemProcessId
                            , e.ProcessId, d.SortNumber, d.ProcessAlias, a.BallMark, a.QmmDetailId, a.Unit
                            , g.MachineName, f.MachineNumber, g.MachineDesc
                            FROM MES.RoutingItemQcItem a
                            INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
							INNER JOIN MES.RoutingItemProcess c ON a.ItemProcessId = c. ItemProcessId
							INNER JOIN MES.RoutingProcess d ON c.RoutingProcessId = d.RoutingProcessId
							INNER JOIN MES.Process e ON d.ProcessId = e.ProcessId
							LEFT JOIN QMS.QmmDetail f ON a.QmmDetailId = f.QmmDetailId
							LEFT JOIN MES.Machine g ON f.MachineId = g.MachineId
                            WHERE a.RoutingItemId = @RoutingItemId AND a.ItemProcessId = @ItemProcessId";
                    dynamicParameters.Add("RoutingItemId", RoutingItemId);
                    dynamicParameters.Add("ItemProcessId", ItemProcessId);


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

        #region //GetLaserEngravedUser //取得雷射刻印人員 --Luca 2025 0306
        public string GetLaserEngravedUser(string Company, string LoginNo, string PassWord)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    sql = @"SELECT a.RoleType,b.UserNo
                            FROM MES.LaserEngravedUser a
                            INNER JOIN BAS.[User] b ON b.UserId=a.UserId
                            INNER JOIN BAS.Department c ON b.DepartmentId=c.DepartmentId
                            INNER JOIN BAS.Company d ON c.CompanyId=d.CompanyId
                            WHERE 1=1 
                            AND a.LoginNo = @LoginNo
                            AND a.PassWord  = @PassWord
                            AND d.CompanyNo  = @Company
                    ";
                    dynamicParameters.Add("LoginNo", LoginNo);
                    dynamicParameters.Add("PassWord", PassWord);
                    dynamicParameters.Add("Company", Company);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("此人員不具備雷刻機使用資格!");

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

        #region //GetLensCarrierRingManagment -- 取得套環基础資料 --Jean 2025.06.30
        public string GetLensCarrierRingManagment(int RingId, string ModelName, string RingCode, string Customer, string Status, string SelStockAlert
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "x.RingId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,x.ModelName,x.RingName,x.Remarks,x.HoleCount,x.RingSpec,x.RingCode,x.RingShape,x.Customer
                            ,x.CreateBy,x.CreateDate,x.RingCodePix,x.Quantity,x.Status,x.DailyDemand,x.SafetyStock
                            ,CASE WHEN x.Quantity<x.SafetyStock THEN N'低于安全库存' ELSE N'安全库存' END AS SafetyStockLevel
                          ";
                    sqlQuery.mainTables =
                        @"FROM (
                            SELECT a.RingId,a.ModelName,a.RingName,a.Remarks,a.HoleCount
                            ,a.RingSpec,a.RingCode,a.RingShape,a.Customer
                            ,d.UserNo+ ' '+d.UserName as CreateBy
                            ,CONVERT(VARCHAR(10), a.CreateDate, 23) as CreateDate
                            ,a.RingCode +' '+a.ModelName as RingCodePix
                            ,ISNULL(SUM(c.Quantity),0)Quantity
                            ,a.Status,CAST(ISNULL(a.DailyDemand,0) AS INT) DailyDemand,CAST(ISNULL(a.SafetyStock,0) AS INT) SafetyStock
                            FROM  MES.LensCarrierRing a
                            LEFT JOIN (SELECT e.RingId,(CASE WHEN e.TransType='IN' THEN e.Quantity WHEN e.TransType='OUT' THEN -e.Quantity ELSE 0 END) Quantity
			                            FROM MES.RingTransaction e
			                            WHERE e.Status='A'
		                            )c ON c.RingId = a.RingId
                            INNER JOIN BAS.[User] d on a.CreateBy = d.UserId
                            WHERE 1=1
                            GROUP BY a.RingId,a.ModelName,a.RingName,a.Remarks,a.HoleCount,a.RingSpec,a.RingCode,a.RingShape,a.Customer,d.UserNo,d.UserName,a.CreateDate,a.Status,a.DailyDemand,a.SafetyStock
                            )x
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @" ";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RingId", @" AND x.RingId = @RingId", RingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModelName", @" AND x.ModelName  LIKE '%' + @ModelName + '%'", ModelName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RingCode", @" AND x.RingCode LIKE '%' + @RingCode + '%'", RingCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Customer", @" AND x.Customer LIKE '%' + @Customer + '%'", Customer);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND x.Status IN @Status", Status.Split(','));

                    switch (SelStockAlert)
                    {
                        case "safe":
                            queryCondition += @" AND x.Quantity >= x.SafetyStock ";
                            break;
                        case "below":
                            queryCondition += @" AND x.Quantity<x.SafetyStock ";
                            break;
                    }

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "x.RingId DESC";
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

        #region //GetLCRFileInfoById -- 取得檔案資料 -- Jean 2025.07.04
        public LCRFileInfo GetLCRFileInfoById(int FileId)
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

                    var result = sqlConnection.Query<LCRFileInfo>(sql, dynamicParameters).FirstOrDefault();

                    return result;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return new LCRFileInfo();
        }
        #endregion

        #region //GetRingTransactionManagment -- 取得套環异动基础資料 --Jean 2025.07.04
        public string GetRingTransactionManagment(int RingTransId, string ModelName, string RingName, string StartTransDate , string EndTransDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.RingTransId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.RingId,a.ModelName,a.RingName
                            ,CASE WHEN a.TransType = 'IN' THEN N'入库' 
	                              WHEN a.TransType = 'OUT' THEN N'出库'
	                              END AS TransTypeName
                            ,a.TransType
                            ,CONVERT(VARCHAR(10), a.TransDate, 23) as TransDate
                            ,a.Quantity
                            ,c.UserNo+ ' '+c.UserName as CreateBy
                            ,CONVERT(VARCHAR(10), a.CreateDate, 23) as CreateDate
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.RingTransaction a
                          INNER JOIN BAS.[User] c on a.CreateBy = c.UserId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @" ";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RingTransId", @" AND a.RingTransId = @RingTransId", RingTransId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModelName", @" AND a.ModelName  LIKE '%' + @ModelName + '%'", ModelName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RingName", @" AND a.RingName LIKE '%' + @RingName + '%'", RingName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartTransDate", @" AND CONVERT(VARCHAR(10), a.TransDate, 120) >= @StartTransDate", StartTransDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndTransDate", @" AND CONVERT(VARCHAR(10), a.TransDate, 120) <= @EndTransDate", EndTransDate);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RingTransId DESC";
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
        #region //AddProdMode -- 生產模式新增 -- Zoey 2022.05.20
        public string AddProdMode(int ModeId, string ModeNo, string ModeName, string ModeDesc, string BarcodeCtrl ,string ScrapRegister
            , string Source, string PVTQCFlag, string NgToBarcode, string TrayBarcode, string LotStatus, string OutputBarcodeFlag, string MrType
            , string OQcCheckType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ModeNo.Length <= 0) throw new SystemException("【模式代碼】不能為空!");
                        if (ModeNo.Length > 50) throw new SystemException("【模式代碼】長度錯誤!");
                        if (ModeName.Length <= 0) throw new SystemException("【模式名稱】不能為空!");
                        if (ModeName.Length > 100) throw new SystemException("【模式名稱】長度錯誤!");
                        if (ModeDesc.Length <= 0) throw new SystemException("【模式描述】不能為空!");
                        if (ModeDesc.Length > 100) throw new SystemException("【模式描述】長度錯誤!");
                        if (BarcodeCtrl.Length <= 0) throw new SystemException("【過站是否要條碼控制】不能為空!");
                        if (ScrapRegister.Length <= 0) throw new SystemException("【報廢品是否入報廢倉】不能為空!");

                        #region //判斷生產模式代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdMode
                                WHERE CompanyId = @CompanyId
                                AND ModeNo = @ModeNo
                                AND ModeId != @ModeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ModeNo", ModeNo);
                        dynamicParameters.Add("ModeId", ModeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【生產模式代碼】重複，請重新輸入!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ProdMode (CompanyId, ModeNo, ModeName, ModeDesc, Status ,BarcodeCtrl, ScrapRegister, OQcCheckType
                                , Source, PVTQCFlag, NgToBarcode, TrayBarcode, LotStatus, OutputBarcodeFlag, MrType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ModeId
                                VALUES (@CompanyId, @ModeNo, @ModeName, @ModeDesc, @Status, @BarcodeCtrl, @ScrapRegister, @OQcCheckType
                                , @Source, @PVTQCFlag, @NgToBarcode, @TrayBarcode, @LotStatus, @OutputBarcodeFlag, @MrType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ModeNo,
                                ModeName,
                                ModeDesc,
                                Status = "A", //啟用
                                BarcodeCtrl,
                                ScrapRegister,
                                OQcCheckType,
                                Source,
                                PVTQCFlag,
                                NgToBarcode,
                                TrayBarcode,
                                LotStatus,
                                OutputBarcodeFlag,
                                MrType,
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

        #region //AddProdModeShift -- 生產模式班次資料新增 -- Shintokru 2022.06.20
        public string AddProdModeShift(int ModeId, int ShiftId, string EffectiveDate, string ExpirationDate)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ModeId < 0) throw new SystemException("【生產模式】不能為空!");
                        if (ShiftId < 0) throw new SystemException("【班次】不能為空!");
                        if (EffectiveDate.Length <= 0) throw new SystemException("【生效日】不能為空!");
                        if (ExpirationDate.Length > 0)
                        {
                            if (EffectiveDate.Equals(ExpirationDate)) throw new SystemException("【生效日】與【失效日】不能相同!");
                            if (DateTime.Compare(Convert.ToDateTime(EffectiveDate), Convert.ToDateTime(ExpirationDate)) > 0) throw new SystemException("【生效日】不能大於【失效日】!");
                        }

                        #region //判斷生產模式資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdMode
                                WHERE CompanyId = @CompanyId
                                AND ModeId = @ModeId
                                AND [Status] = 'A'";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ModeId", ModeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式資料錯誤!");
                        #endregion

                        #region //判斷班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Shift]
                                WHERE CompanyId = @CompanyId
                                AND ShiftId = @ShiftId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("班次資料錯誤!");
                        #endregion

                        #region //判斷生產模式班次資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdModeShift
                                WHERE ModeId = @ModeId
                                AND ShiftId = @ShiftId";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【生產模式班次】重複，請重新輸入!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ProdModeShift (ModeId, ShiftId, EffectiveDate, ExpirationDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ModeShiftId
                                VALUES (@ModeId, @ShiftId, @EffectiveDate, @ExpirationDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModeId,
                                ShiftId,
                                EffectiveDate,
                                ExpirationDate = ExpirationDate.Length > 0 ? ExpirationDate : null,
                                Status = "A", //啟用
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

        #region //AddWorkShop -- 車間資料新增 -- Shintokru 2022.06.07
        public string AddWorkShop(string ShopNo, string ShopName, string ShopDesc, string Location, string Floor)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ShopNo.Length <= 0) throw new SystemException("【車間編號】不能為空!");
                        if (ShopNo.Length > 10) throw new SystemException("【車間編號】長度錯誤!");
                        if (ShopName.Length <= 0) throw new SystemException("【車間名稱】不能為空!");
                        if (ShopName.Length > 100) throw new SystemException("【車間名稱】長度錯誤!");
                        if (ShopDesc.Length <= 0) throw new SystemException("【車間描述】不能為空!");
                        if (ShopDesc.Length > 100) throw new SystemException("【車間描述】長度錯誤!");
                        if (Location.Length > 100) throw new SystemException("【位置】長度錯誤!");
                        if (Floor.Length > 100) throw new SystemException("【樓層】長度錯誤!");

                        #region //判斷車間代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopNo = @ShopNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopNo", ShopNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【車間代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷車間名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopName = @ShopName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopName", ShopName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【車間名稱】重複，請重新輸入!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.WorkShop (CompanyId, ShopNo, ShopName, ShopDesc, Location, Floor, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ShopId
                                VALUES (@CompanyId, @ShopNo, @ShopName, @ShopDesc, @Location, @Floor, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ShopNo,
                                ShopName,
                                ShopDesc,
                                Location = Location.Length > 0 ? Location :null,
                                Floor = Floor.Length > 0 ? Floor : null,
                                Status = "A", //啟用
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

        #region //AddMachine -- 機台資料新增 -- Ted 2022.06.08
        public string AddMachine(int ShopId, string MachineNo, string MachineName, string MachineDesc)
        {
            try
            {
                if (ShopId <= 0) throw new SystemException("【所屬車間】不能為空!");
                if (MachineNo.Length <= 0) throw new SystemException("【機台編號】不能為空!");
                if (MachineNo.Length > 10) throw new SystemException("【機台編號】長度錯誤!");
                if (MachineName.Length <= 0) throw new SystemException("【機台名稱】不能為空!");
                if (MachineName.Length > 100) throw new SystemException("【機台名稱】長度錯誤!");
                if (MachineDesc.Length <= 0) throw new SystemException("【機台描述】不能為空!");
                if (MachineDesc.Length > 100) throw new SystemException("【機台描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷車間資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopId = @ShopId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopId", ShopId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");
                        #endregion

                        #region //判斷機台代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine a
                                INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId
                                WHERE b.CompanyId = @CompanyId
                                AND a.MachineNo = @MachineNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MachineNo", MachineNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【機台代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Machine (ShopId, MachineNo, MachineName, MachineDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MachineId
                                VALUES (@ShopId, @MachineNo, @MachineName, @MachineDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShopId,
                                MachineNo,
                                MachineName,
                                MachineDesc,
                                Status = "A", //啟用
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

        #region //AddMachineAsset --機台資產資料新增 -- Ted 2022.06.10
        public string AddMachineAsset(int MachineId, string AssetNo)
        {
            try
            {
                if (AssetNo.Length <= 0) throw new SystemException("【機台資產編號】不能為空!");
                if (AssetNo.Length > 30) throw new SystemException("【機台資產編號】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine a
                                INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId
                                WHERE a.MachineId = @MachineId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        #region //判機台資產資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineAsset
                                WHERE MachineId = @MachineId
                                AND AssetNo = @AssetNo";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("AssetNo", AssetNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【機台資產】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MachineAsset (MachineId, AssetNo
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@MachineId, @AssetNo
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MachineId,
                                AssetNo,
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

        #region //AddTimeInterval -- 時間區段新增 -- Shintokru 2022.06.15
        public string AddTimeInterval(string BeginTime, string EndTime)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (BeginTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                        if (EndTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");

                        #region //判斷開始時間是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.TimeInterval
                                WHERE CompanyId = @CompanyId
                                AND BeginTime = @BeginTime";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("BeginTime", BeginTime);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【開始時間】重複，請重新輸入!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.TimeInterval (CompanyId, BeginTime, EndTime
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.IntervalId
                                VALUES (@CompanyId, @BeginTime, @EndTime
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                BeginTime,
                                EndTime,
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

        #region //AddDevice -- 裝置資料新增 -- Ted 2022.06.16
        public string AddDevice(string DeviceType, string DeviceName, string DeviceDesc, string DeviceIdentifierCode, string DeviceAuthority)
        {
            try
            {
                if (DeviceType.Length <= 0) throw new SystemException("【設備類別】不能為空!");
                if (DeviceName.Length <= 0) throw new SystemException("【設備識別碼】不能為空!");
                if (DeviceName.Length > 100) throw new SystemException("【設備識別碼】長度錯誤!");
                if (DeviceDesc.Length <= 0) throw new SystemException("【設備識別碼】不能為空!");
                if (DeviceDesc.Length > 100) throw new SystemException("【設備識別碼】長度錯誤!");
                if (DeviceIdentifierCode.Length <= 0) throw new SystemException("【設備識別碼】不能為空!");
                if (DeviceIdentifierCode.Length > 50) throw new SystemException("【設備識別碼】長度錯誤!");
                if (DeviceAuthority.Length <= 0) throw new SystemException("【設備權限】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷設備名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceName = @DeviceName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceName", DeviceName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【設備名稱】重複，請重新輸入!");
                        #endregion

                        #region //判斷設備識別碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceIdentifierCode = @DeviceIdentifierCode";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceIdentifierCode", DeviceIdentifierCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【設備識別碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Device (CompanyId, DeviceType, DeviceName, DeviceDesc, DeviceIdentifierCode, DeviceAuthority, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DeviceId
                                VALUES (@CompanyId, @DeviceType, @DeviceName, @DeviceDesc, @DeviceIdentifierCode, @DeviceAuthority, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                DeviceType,
                                DeviceName,
                                DeviceDesc,
                                DeviceIdentifierCode,
                                DeviceAuthority,
                                Status = "A", //啟用
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

        #region //AddDeviceMachine -- 裝置機台綁定 -- Ted 2022.06.16
        public string AddDeviceMachine(int DeviceId, int MachineId)
        {
            try
            { 
                if (DeviceId <= 0) throw new SystemException("【裝置】不能為空!");
                if (MachineId <= 0) throw new SystemException("【機台】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.DeviceMachine
                                WHERE DeviceId = @DeviceId";
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【裝置】錯誤，請重新輸入!");
                        #endregion

                        #region //判斷機台是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.DeviceMachine
                                WHERE 1=1
                                AND DeviceId = @DeviceId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("DeviceId", DeviceId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【機台】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.DeviceMachine (DeviceId, MachineId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@DeviceId, @MachineId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DeviceId,
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

        #region //AddProcess -- 製程資料新增 -- Shintokru 2022.06.20
        public string AddProcess(string ProcessNo, string ProcessName, string ProcessDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ProcessNo.Length <= 0) throw new SystemException("【製程代碼】不能為空!");
                        if (ProcessNo.Length > 20) throw new SystemException("【製程代碼】長度錯誤!");
                        if (ProcessName.Length <= 0) throw new SystemException("【製程名稱】不能為空!");
                        if (ProcessName.Length > 100) throw new SystemException("【製程名稱】長度錯誤!");
                        if (ProcessDesc.Length <= 0) throw new SystemException("【製程描述】不能為空!");
                        if (ProcessDesc.Length > 100) throw new SystemException("【製程描述】長度錯誤!");

                        #region //判斷製程代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE CompanyId = @CompanyId
                                AND ProcessNo = @ProcessNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ProcessNo", ProcessNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【製程代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷製程名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE CompanyId = @CompanyId
                                AND ProcessName = @ProcessName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ProcessName", ProcessName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【製程名稱】重複，請重新輸入!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Process (CompanyId, ProcessNo, ProcessName, ProcessDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProcessId
                                VALUES (@CompanyId, @ProcessNo, @ProcessName, @ProcessDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ProcessNo,
                                ProcessName,
                                ProcessDesc,
                                Status = "A", //啟用
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

        #region //AddProdUnit -- 生產單元資料新增 -- Ted 2022.06.23
        public string AddProdUnit(string UnitNo, string UnitName, string UnitDesc, string CheckStatus)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (UnitNo.Length <= 0) throw new SystemException("【單元代碼】不能為空!");
                        if (UnitNo.Length > 10) throw new SystemException("【單元代碼】長度錯誤!");
                        if (UnitName.Length <= 0) throw new SystemException("【單元名稱】不能為空!");
                        if (UnitName.Length > 100) throw new SystemException("【單元名稱】長度錯誤!");
                        if (UnitDesc.Length <= 0) throw new SystemException("【單元描述】不能為空!");
                        if (UnitDesc.Length > 100) throw new SystemException("【單元描述】長度錯誤!");

                        #region //判斷單元代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdUnit
                                WHERE CompanyId = @CompanyId
                                AND UnitNo = @UnitNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UnitNo", UnitNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【單元代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷單元名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdUnit
                                WHERE CompanyId = @CompanyId
                                AND UnitName = @UnitName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UnitName", UnitName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【單元名稱】重複，請重新輸入!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ProdUnit (CompanyId, UnitNo, UnitName, UnitDesc, CheckStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UnitId
                                VALUES (@CompanyId, @UnitNo, @UnitName, @UnitDesc, @CheckStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                UnitNo,
                                UnitName,
                                UnitDesc,
                                CheckStatus,
                                Status = "A", //啟用
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

        #region //AddWarehouse -- 庫房基本資料新增 -- Ted 2022.07.01
        public string AddWarehouse(int ShopId, string WarehouseNo, string WarehouseName, string WarehouseDesc)
        {
            try
            {
                if (ShopId <= 0) throw new SystemException("【車間代碼】不能為空!");
                if (WarehouseNo.Length <= 0) throw new SystemException("【庫房代碼】不能為空!");
                if (WarehouseNo.Length > 10) throw new SystemException("【庫房代碼】長度錯誤!");
                if (WarehouseName.Length <= 0) throw new SystemException("【庫房名稱】不能為空!");
                if (WarehouseName.Length > 100) throw new SystemException("【庫房名稱】長度錯誤!");
                if (WarehouseDesc.Length <= 0) throw new SystemException("【庫房描述】不能為空!");
                if (WarehouseDesc.Length > 100) throw new SystemException("【庫房描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷庫房代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Warehouse a
                                INNER JOIN MES.WorkShop b on a.ShopId = b.ShopId
                                WHERE b.CompanyId = @CompanyId
                                AND a.WarehouseNo = @WarehouseNo
                                AND a.ShopId != @ShopId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("WarehouseNo", WarehouseNo);
                        dynamicParameters.Add("ShopId", ShopId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【庫房代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.Warehouse (ShopId, WarehouseNo, WarehouseName, WarehouseDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.WarehouseId
                                VALUES (@ShopId, @WarehouseNo, @WarehouseName, @WarehouseDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ShopId,
                                WarehouseNo,
                                WarehouseName,
                                WarehouseDesc,
                                Status = "A", //啟用
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

        #region //AddWarehouseLocation -- 庫房儲位基本資料新增 -- Ted 2022.07.04
        public string AddWarehouseLocation(int WarehouseId, string LocationNo, string LocationName, string LocationDesc)
        {
            try
            {
                if (WarehouseId <= 0) throw new SystemException("【庫房代碼】不能為空!");
                if (LocationNo.Length <= 0) throw new SystemException("【儲位代碼】不能為空!");
                if (LocationNo.Length > 10) throw new SystemException("【儲位代碼】長度錯誤!");
                if (LocationName.Length <= 0) throw new SystemException("【儲位名稱】不能為空!");
                if (LocationName.Length > 100) throw new SystemException("【儲位名稱】長度錯誤!");
                if (LocationDesc.Length <= 0) throw new SystemException("【儲位描述】不能為空!");
                if (LocationDesc.Length > 100) throw new SystemException("【儲位描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷儲位代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WarehouseLocation a
                                INNER JOIN MES.Warehouse b on a.WarehouseId = b.WarehouseId
                                INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                                WHERE c.CompanyId = @CompanyId
                                AND a.LocationNo = @LocationNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("LocationNo", LocationNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【儲位代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.WarehouseLocation (WarehouseId, LocationNo, LocationName, LocationDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.LocationId
                                VALUES (@WarehouseId, @LocationNo, @LocationName, @LocationDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                WarehouseId,
                                LocationNo,
                                LocationName,
                                LocationDesc,
                                Status = "A", //啟用
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

        #region //AddProcessParameter -- 製程參數資料新增 -- Ted 2022.07.05
        public string AddProcessParameter(int ProcessId, int ModeId, int DepartmentId, string ProcessCheckStatus, string PreCollectionStatus, string PackageFlag
            , string PostCollectionStatus, string NgToBarcode, string PassingMode, string ProcessCheckType, string ConsumeFlag)
        {
            try
            {
                if (ProcessId <= 0) throw new SystemException("【製程】不能為空!");
                if (ModeId <= 0) throw new SystemException("【生產模式】不能為空!");
                if (ProcessCheckStatus.Length > 0)
                {
                    if (ProcessCheckStatus == "Y")
                    {
                        if (ProcessCheckType.Length <= 0) throw new SystemException("【工程檢頻率】不能為空!");
                    }
                }
                if (PreCollectionStatus.Length <= 0) throw new SystemException("【是否收集開工前資訊】不能為空!");
                if (PostCollectionStatus.Length <= 0) throw new SystemException("【是否收集完工後資訊】不能為空!");
                if (NgToBarcode.Length <= 0) throw new SystemException("【是否刷取NG條碼】不能為空!");
                if (PassingMode.Length <= 0) throw new SystemException("【過站模式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷該製程下生產模式是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessParameter
                                WHERE ProcessId = @ProcessId
                                AND ModeId = @ModeId
                                AND DepartmentId = @DepartmentId";
                        dynamicParameters.Add("ProcessId", ProcessId);
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【該製程下生產模式】重複，請重新輸入!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ProcessParameter (ProcessId, ModeId, DepartmentId, ProcessCheckStatus, PreCollectionStatus
                                , PostCollectionStatus, NgToBarcode, PassingMode, ProcessCheckType ,Status, PackageFlag, ConsumeFlag
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ParameterId
                                VALUES (@ProcessId, @ModeId, @DepartmentId, @ProcessCheckStatus, @PreCollectionStatus
                                , @PostCollectionStatus, @NgToBarcode, @PassingMode, @ProcessCheckType, @Status, @PackageFlag, @ConsumeFlag
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessId,
                                ModeId,
                                DepartmentId,
                                ProcessCheckStatus,
                                PreCollectionStatus,
                                PostCollectionStatus,
                                NgToBarcode,
                                PassingMode,
                                ProcessCheckType,
                                Status = "A", //啟用
                                PackageFlag,
                                ConsumeFlag,
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

        #region //AddProcessMachine -- 新增製程機台資料 -- Shintokru 2022.07.06
        public string AddProcessMachine(int ParameterId, int MachineId, int ToolCount, string KeyenceFlag, int KeyenceId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ParameterId <= 0) throw new SystemException("【製程】不能為空!");
                        if (MachineId <= 0) throw new SystemException("【機台】不能為空!");
                        if (ToolCount <= 0) throw new SystemException("【工具上限數】至少為1!");
                        if (KeyenceFlag == "Y" && KeyenceId <= 0) throw new SystemException("KeyenceID不能為空!");

                        #region //判斷該製程下機台是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessMachine
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機台】重複，請重新輸入!");
                        #endregion

                        if (KeyenceId > 0)
                        {
                            #region //確認Keyence設備資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Keyence a
                                    WHERE a.KeyenceId = @KeyenceId";
                            dynamicParameters.Add("KeyenceId", KeyenceId);

                            var KeyenceResult = sqlConnection.Query(sql, dynamicParameters);

                            if (KeyenceResult.Count() <= 0) throw new SystemException("Keyence設備資料錯誤!!");
                            #endregion
                        }

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ProcessMachine (ParameterId, MachineId, BatchStatus, ToolCount, KeyenceFlag, KeyenceId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@ParameterId, @MachineId, @BatchStatus, @ToolCount, @KeyenceFlag, @KeyenceId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ParameterId,
                                MachineId,
                                BatchStatus = "Y",
                                ToolCount,
                                KeyenceFlag,
                                KeyenceId = KeyenceId > 0 ? KeyenceId : (int?)null,
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

        #region //AddProcessProductionUnit -- 新增製程生產單元資料 -- Ted 2022.07.06
        public string AddProcessProductionUnit(int ParameterId, int UnitId, int SortNumber)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ParameterId <= 0) throw new SystemException("【製程】不能為空!");
                        if (UnitId <= 0) throw new SystemException("【機台】不能為空!");
                        if (SortNumber <= 0) throw new SystemException("【排序號碼】不能為空!");

                        #region //判斷該製程下單元是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessProductionUnit
                                WHERE ParameterId = @ParameterId
                                AND UnitId = @UnitId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("UnitId", UnitId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【單元】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.ProcessProductionUnit (ParameterId, UnitId, SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@ParameterId, @UnitId, @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ParameterId,
                                UnitId,
                                SortNumber,
                                Status = "A", //啟用
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

        #region //AddTray -- 托盤資料新增 -- Ted 2023.02.04
        public string AddTray(string TrayPrefix, string TrayName, int TrayCapacity,int SerialNumber, int Fabrication, int Serial, string SuffixCode
            , string Remark, int ViewCompanyId)
        {
            try
            {
                if (TrayPrefix.Length <= 0) throw new SystemException("【托盤前綴】不能為空!");
                if (TrayPrefix.Length > 100) throw new SystemException("【托盤前綴】長度錯誤!");
                if (TrayName.Length <= 0) throw new SystemException("【托盤名稱】不能為空!");
                if (TrayName.Length > 100) throw new SystemException("【托盤名稱】長度錯誤!");
                if (TrayCapacity < 0) throw new SystemException("【托盤容量】不能為小於0!");
                if (Serial <= 0) throw new SystemException("【流水碼數】不能為小於0!");
                if(SerialNumber > 0)
                {
                    if (Fabrication != 0) throw new SystemException("【流水編號模式】不須Key欲生產托盤數量!");
                }
                if (Fabrication > 0)
                {
                    if (SerialNumber != 0) throw new SystemException("【欲生產托盤數量模式】不須Key流水編號!");
                    if (Fabrication > 5000) throw new SystemException("【欲生產托盤數量】一次生產最多不能超過5000,如需要請分批生產");
                }
                string NotAccomplished = "";
                int? nullData = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        int serialNumMax = 0;
                        string SerialNo = "";
                        string SerialZero = "";
                        for (var i = 1; i <= Serial; i++)
                        {
                            SerialNo += "_";
                            SerialZero += "0";
                        }
                        string pattern = @"^[A-Z0-9\\-]*$";
                        if (!Regex.Match(TrayPrefix, pattern).Success)
                        {
                            throw new SystemException("【托盤管理】- 托盤編號只能由大寫英文和數字和【-】組成");
                        }

                        if (SuffixCode == "N")
                        {
                            #region //判斷托盤編號最大值
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TrayNo 
                                FROM MES.Tray a
                                WHERE a.TrayNo LIKE @TrayPrefix + @SerialNo
                                ORDER BY a.TrayId DESC";
                            dynamicParameters.Add("TrayPrefix", TrayPrefix);
                            dynamicParameters.Add("SerialNo", SerialNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() > 0)
                            {
                                string TrayNoBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TrayNo;
                                if (TrayNoBase != null)
                                {
                                    //TrayNoBase = "Tray-0001";
                                    //string[] splitStr = { TrayPrefix }; //自行設定切割字串

                                    serialNumMax = Convert.ToInt32(TrayNoBase.Substring(TrayNoBase.Length - Serial));
                                }
                            }
                            #endregion
                        }
                        else if (SuffixCode == "Y")
                        {
                            #region //判斷托盤編號最大值
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TrayNo 
                                FROM MES.Tray a
                                WHERE a.TrayNo LIKE @TrayPrefix + '-' + @SerialNo
                                ORDER BY a.TrayId DESC";
                            dynamicParameters.Add("TrayPrefix", TrayPrefix);
                            dynamicParameters.Add("SerialNo", SerialNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() > 0)
                            {
                                string TrayNoBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TrayNo;
                                if (TrayNoBase != null)
                                {
                                    //TrayNoBase = "Tray-0001";
                                    //string[] splitStr = { TrayPrefix }; //自行設定切割字串

                                    serialNumMax = Convert.ToInt32(TrayNoBase.Substring(TrayNoBase.Length - Serial));
                                }
                            }
                            #endregion
                        }



                        #region //判斷托盤編號是否重複
                        int rowsAffected = 0;
                        int ProdtuctNum = 0;
                        if (Fabrication > 0)
                        {
                            //生產數量模式
                            ProdtuctNum = serialNumMax + Fabrication;
                        }
                        if (SerialNumber > 0)
                        {
                            //流水編號模式
                            if (serialNumMax >= SerialNumber) throw new SystemException("【托盤管理】- 目前設定流水編號系統已經存在!");
                            ProdtuctNum = SerialNumber;
                        }

                        int NowNum = 1;
                        for (var i = serialNumMax + 1; i <= ProdtuctNum; i++)
                        {
                            
                            
                            string serialNum = String.Format("{0:" + SerialZero + "}", Convert.ToInt16(i));
                            string TrayNo = "";
                            if (SuffixCode == "Y")
                            {
                                TrayNo = TrayPrefix + '-' + serialNum;

                            }
                            else if (SuffixCode == "N")
                            {
                                TrayNo = TrayPrefix + serialNum;
                            }
                            if (SerialZero.Length != serialNum.Length) throw new SystemException("【托盤管理】- 已達到該組合流水碼最大號了!");
                            
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Tray
                                    WHERE TrayNo = @TrayNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("TrayNo", TrayNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】- 托盤編號重複，請重新輸入!");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Barcode
                                    WHERE BarcodeNo = @TrayNo";
                            dynamicParameters.Add("TrayNo", TrayNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】- 托盤編號已存在Barcode，請重新輸入!");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.BarcodePrint
                                    WHERE BarcodeNo = @TrayNo";
                            dynamicParameters.Add("TrayNo", TrayNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】- 托盤編號已存在BarcodePrint，請重新輸入!");

                            #region //資料新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.Tray (CompanyId, BarcodeNo, TrayNo, TrayName, TrayCapacity, Remark, UseTimes, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TrayId
                                VALUES (@CompanyId, @BarcodeNo, @TrayNo, @TrayName, @TrayCapacity, @Remark, @UseTimes, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    BarcodeNo = nullData,
                                    TrayNo,
                                    TrayName,
                                    TrayCapacity,
                                    Remark,
                                    UseTimes = '0',
                                    Status = "A", //啟用
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

                            if (SerialNumber > 0)
                            {
                                NowNum++;

                                if (NowNum > 5000 && ProdtuctNum != i)
                                {
                                    NotAccomplished = "不能一次性產生超過5000筆資料,目前Tray編號產到" + TrayNo + ",請接續新增";
                                    break;
                                }
                            }

                        }
                        #endregion


                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = NotAccomplished,
                            msg = "(" + rowsAffected + " rows affected)",
                        });
                        #endregion

                        if (ViewCompanyId != CurrentCompany) throw new SystemException("頁面公司別與系統後台公司別不相同,請重新登入系統");
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

        #region//apiAddTray  -- 托盤資料新增 -- Luca 2024.12.23
        public string apiAddTray(string Company,string UserNo ,string TrayPrefix, string TrayName, int TrayCapacity, int SerialNumber, int Fabrication, int Serial, string SuffixCode
                    , string Remark)
        {
            try
            {
                if (TrayPrefix.Length <= 0) throw new SystemException("【托盤前綴】不能為空!");
                if (TrayPrefix.Length > 100) throw new SystemException("【托盤前綴】長度錯誤!");
                if (TrayName.Length <= 0) throw new SystemException("【托盤名稱】不能為空!");
                if (TrayName.Length > 100) throw new SystemException("【托盤名稱】長度錯誤!");
                if (TrayCapacity < 0) throw new SystemException("【托盤容量】不能為小於0!");
                if (Serial <= 0) throw new SystemException("【流水碼數】不能為小於0!");
                if (SerialNumber > 0)
                {
                    if (Fabrication != 0) throw new SystemException("【流水編號模式】不須Key欲生產托盤數量!");
                }
                if (Fabrication > 0)
                {
                    if (SerialNumber != 0) throw new SystemException("【欲生產托盤數量模式】不須Key流水編號!");
                    if (Fabrication > 5000) throw new SystemException("【欲生產托盤數量】一次生產最多不能超過5000,如需要請分批生產");
                }
           
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        DynamicParameters dynamicParameters = new DynamicParameters();
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

                        #region //取得User資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                        int UserId = -1;
                        foreach (var item in UserResult)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        int serialNumMax = 0;
                        string SerialNo = "";
                        string SerialZero = "";
                        for (var i = 1; i <= Serial; i++)
                        {
                            SerialNo += "_";
                            SerialZero += "0";
                        }
                        string pattern = @"^[A-Z0-9\\-]*$";
                        if (!Regex.Match(TrayPrefix, pattern).Success)
                        {
                            throw new SystemException("【托盤管理】- 托盤編號只能由大寫英文和數字和【-】組成");
                        }

                        if (SuffixCode == "N")
                        {
                            //同時檢查 MES.Tray 和 MES.TrayTemp 兩個表格，然後取得兩者中較大的 TrayNo
                            #region //判斷托盤編號最大值
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TrayNo
                                    FROM (
                                        SELECT TrayNo, TrayId 
                                        FROM MES.Tray 
                                        WHERE TrayNo LIKE @TrayPrefix + @SerialNo
                                        UNION ALL
                                        SELECT a.TrayTempNo TrayNo,a.TrayTempId TrayId
                                        FROM MES.TrayTemp a
                                        WHERE a.TrayTempNo LIKE @TrayPrefix + @SerialNo
                                    ) AS CombinedTray 
                                    ORDER BY TrayNo DESC";
                            dynamicParameters.Add("TrayPrefix", TrayPrefix);
                            dynamicParameters.Add("SerialNo", SerialNo);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                string TrayNoBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TrayNo;
                                if (TrayNoBase != null)
                                {
                                    serialNumMax = Convert.ToInt32(TrayNoBase.Substring(TrayNoBase.Length - Serial));
                                }
                            }
                            #endregion
                        }
                        else if (SuffixCode == "Y")
                        {                            
                            #region //判斷托盤編號最大值
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TrayNo
                            FROM (
                                SELECT TrayNo, TrayId 
                                FROM MES.Tray 
                                WHERE TrayNo LIKE @TrayPrefix + @SerialNo
                                UNION ALL
                                SELECT a.TrayTempNo TrayNo,a.TrayTempId TrayId
                                FROM MES.TrayTemp a
                                WHERE a.TrayTempNo LIKE @TrayPrefix + @SerialNo
                            ) AS CombinedTray 
                            ORDER BY TrayNo DESC";
                            dynamicParameters.Add("TrayPrefix", TrayPrefix);
                            dynamicParameters.Add("SerialNo", SerialNo);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                string TrayNoBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TrayNo;
                                if (TrayNoBase != null)
                                {
                                    serialNumMax = Convert.ToInt32(TrayNoBase.Substring(TrayNoBase.Length - Serial));
                                }
                            }                            
                            #endregion
                        }

                        #region //判斷托盤編號是否重複
                        int rowsAffected = 0;
                        int ProdtuctNum = 0;
                        if (Fabrication > 0)
                        {
                            //生產數量模式
                            ProdtuctNum = serialNumMax + Fabrication;
                        }
                        if (SerialNumber > 0)
                        {
                            //流水編號模式
                            if (serialNumMax >= SerialNumber) throw new SystemException("【托盤管理】- 目前設定流水編號系統已經存在!");
                            ProdtuctNum = SerialNumber;
                        }

                        int NowNum = 1;                        
                        for (var i = serialNumMax + 1; i <= ProdtuctNum; i++)
                        {
                            string serialNum = String.Format("{0:" + SerialZero + "}", Convert.ToInt16(i));
                            string TrayNo = "";
                            if (SuffixCode == "Y")
                            {
                                TrayNo = TrayPrefix + '-' + serialNum;

                            }
                            else if (SuffixCode == "N")
                            {
                                TrayNo = TrayPrefix + serialNum;
                            }
                            if (SerialZero.Length != serialNum.Length) throw new SystemException("【托盤管理】- 已達到該組合流水碼最大號了!");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Tray
                                    WHERE TrayNo = @TrayNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("TrayNo", TrayNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】- 托盤編號重複，請重新輸入!");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Barcode
                                    WHERE BarcodeNo = @TrayNo";
                            dynamicParameters.Add("TrayNo", TrayNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】- 托盤編號已存在Barcode，請重新輸入!");

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.BarcodePrint
                                    WHERE BarcodeNo = @TrayNo";
                            dynamicParameters.Add("TrayNo", TrayNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】- 托盤編號已存在BarcodePrint，請重新輸入!");

                            #region //資料新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.TrayTemp (CompanyId, TrayTempNo, TrayTempName, TrayTempCapacity, Remark, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.TrayTempNo AS TrayNo
                                VALUES (@CompanyId, @TrayNo, @TrayName, @TrayCapacity, @Remark, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,                                
                                    TrayNo,
                                    TrayName,
                                    TrayCapacity,
                                    Remark,                               
                                    Status = "A", //啟用
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = insertResult,
                                msg = "(" + rowsAffected + " rows affected)",
                            });
                            #endregion

                        }
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

        #region// AI建議的apiAddTray托盤新增
        public string apiAddTrayXXX(string Company, string UserNo, string TrayPrefix, string TrayName, int TrayCapacity, int SerialNumber, int Fabrication, int Serial, string SuffixCode, string Remark)
        {
            try
            {
                // 參數驗證
                if (TrayPrefix.Length <= 0) throw new SystemException("【托盤前綴】不能為空!");
                if (TrayPrefix.Length > 100) throw new SystemException("【托盤前綴】長度錯誤!");
                if (TrayName.Length <= 0) throw new SystemException("【托盤名稱】不能為空!");
                if (TrayName.Length > 100) throw new SystemException("【托盤名稱】長度錯誤!");
                if (TrayCapacity < 0) throw new SystemException("【托盤容量】不能為小於0!");
                if (Serial <= 0) throw new SystemException("【流水碼數】不能為小於0!");
                if (SerialNumber > 0)
                {
                    if (Fabrication != 0) throw new SystemException("【流水編號模式】不須Key欲生產托盤數量!");
                }
                if (Fabrication > 0)
                {
                    if (SerialNumber != 0) throw new SystemException("【欲生產托盤數量模式】不須Key流水編號!");
                    if (Fabrication > 5000) throw new SystemException("【欲生產托盤數量】一次生產最多不能超過5000,如需要請分批生產");
                }

                // 正則表達式驗證
                string pattern = @"^[A-Z0-9\\-]*$";
                if (!Regex.Match(TrayPrefix, pattern).Success)
                {
                    throw new SystemException("【托盤管理】- 托盤編號只能由大寫英文和數字和【-】組成");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        sqlConnection.Open();

                        #region //確認公司別DB
                        DynamicParameters dynamicParameters = new DynamicParameters();
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

                        #region //取得User資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId FROM BAS.[User] WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);
                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                        int UserId = UserResult.First().UserId;
                        #endregion

                        // 計算序列號相關參數
                        int serialNumMax = 0;
                        string SerialNo = new string('_', Serial);
                        string SerialZero = new string('0', Serial);

                        #region //取得最大序列號 - 優化查詢
                        dynamicParameters = new DynamicParameters();
                        sql = @"
                            SELECT TOP 1 
                                CASE 
                                    WHEN @SuffixCode = 'Y' THEN 
                                        CAST(RIGHT(TrayNo, @Serial) AS INT)
                                    ELSE 
                                        CAST(RIGHT(TrayNo, @Serial) AS INT)
                                END AS MaxSerial
                            FROM (
                                SELECT TrayNo FROM MES.Tray WITH(NOLOCK)
                                WHERE TrayNo LIKE @TrayPattern
                                AND LEN(TrayNo) = @ExpectedLength
                                UNION ALL
                                SELECT TrayTempNo AS TrayNo FROM MES.TrayTemp WITH(NOLOCK)
                                WHERE TrayTempNo LIKE @TrayPattern
                                AND LEN(TrayTempNo) = @ExpectedLength
                            ) AS CombinedTray 
                            WHERE ISNUMERIC(RIGHT(TrayNo, @Serial)) = 1
                            ORDER BY CAST(RIGHT(TrayNo, @Serial) AS INT) DESC";

                        string trayPattern = SuffixCode == "Y" ? TrayPrefix + "-%" : TrayPrefix + "%";
                        int expectedLength = SuffixCode == "Y" ? TrayPrefix.Length + 1 + Serial : TrayPrefix.Length + Serial;

                        dynamicParameters.Add("TrayPattern", trayPattern);
                        dynamicParameters.Add("Serial", Serial);
                        dynamicParameters.Add("SuffixCode", SuffixCode);
                        dynamicParameters.Add("ExpectedLength", expectedLength);

                        var maxSerialResult = sqlConnection.Query(sql, dynamicParameters);
                        if (maxSerialResult.Count() > 0)
                        {
                            var firstResult = maxSerialResult.First();
                            if (firstResult.MaxSerial != null)
                            {
                                serialNumMax = Convert.ToInt32(firstResult.MaxSerial);
                            }
                        }
                        #endregion

                        #region //計算要生產的托盤數量範圍
                        int startNum = serialNumMax + 1;
                        int endNum = 0;

                        if (Fabrication > 0)
                        {
                            // 生產數量模式
                            endNum = serialNumMax + Fabrication;
                        }
                        else if (SerialNumber > 0)
                        {
                            // 流水編號模式
                            if (serialNumMax >= SerialNumber)
                                throw new SystemException("【托盤管理】- 目前設定流水編號系統已經存在!");
                            endNum = SerialNumber;
                        }

                        // 生成所有要新增的托盤編號
                        List<string> trayNumbers = new List<string>();
                        for (int i = startNum; i <= endNum; i++)
                        {
                            string serialNum = i.ToString(SerialZero);
                            if (SerialZero.Length != serialNum.Length)
                                throw new SystemException("【托盤管理】- 已達到該組合流水碼最大號了!");

                            string trayNo = SuffixCode == "Y" ? $"{TrayPrefix}-{serialNum}" : $"{TrayPrefix}{serialNum}";
                            trayNumbers.Add(trayNo);
                        }
                        #endregion

                        #region //批次檢查重複性 - 大幅優化性能
                        if (trayNumbers.Count > 0)
                        {
                            // 使用 Table-Valued Parameter 或 IN 子句進行批次檢查
                            string trayNoList = string.Join("','", trayNumbers);

                            // 檢查 MES.Tray 重複
                            sql = $@"
                                SELECT TrayNo FROM MES.Tray WITH(NOLOCK)
                                WHERE TrayNo IN ('{trayNoList}')";
                            var duplicateTray = sqlConnection.Query(sql);
                            if (duplicateTray.Count() > 0)
                            {
                                string duplicates = string.Join(", ", duplicateTray.Select(x => x.TrayNo));
                                throw new SystemException($"【托盤管理】- 托盤編號重複: {duplicates}");
                            }

                            // 檢查 MES.Barcode 重複
                            sql = $@"
                                SELECT BarcodeNo FROM MES.Barcode WITH(NOLOCK)
                                WHERE BarcodeNo IN ('{trayNoList}')";
                            var duplicateBarcode = sqlConnection.Query(sql);
                            if (duplicateBarcode.Count() > 0)
                            {
                                string duplicates = string.Join(", ", duplicateBarcode.Select(x => x.BarcodeNo));
                                throw new SystemException($"【托盤管理】- 托盤編號已存在Barcode: {duplicates}");
                            }

                            // 檢查 MES.BarcodePrint 重複
                            sql = $@"
                            SELECT BarcodeNo FROM MES.BarcodePrint WITH(NOLOCK)
                            WHERE BarcodeNo IN ('{trayNoList}')";
                            var duplicateBarcodePrint = sqlConnection.Query(sql);
                            if (duplicateBarcodePrint.Count() > 0)
                            {
                                string duplicates = string.Join(", ", duplicateBarcodePrint.Select(x => x.BarcodeNo));
                                throw new SystemException($"【托盤管理】- 托盤編號已存在BarcodePrint: {duplicates}");
                            }
                        }
                        #endregion

                        #region //插入資料
                        int rowsAffected = 0;
                        List<dynamic> insertResults = new List<dynamic>();
                        DateTime now = DateTime.Now;

                        // 逐筆插入（針對少量資料優化）
                        foreach (string trayNo in trayNumbers)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"
                        INSERT INTO MES.TrayTemp (CompanyId, TrayTempNo, TrayTempName, TrayTempCapacity, Remark, Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                        OUTPUT INSERTED.TrayTempNo AS TrayNo
                        VALUES (@CompanyId, @TrayTempNo, @TrayTempName, @TrayTempCapacity, @Remark, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                            dynamicParameters.AddDynamicParams(new
                            {
                                CompanyId = CurrentCompany,
                                TrayTempNo = trayNo,
                                TrayTempName = TrayName,
                                TrayTempCapacity = TrayCapacity,
                                Remark = Remark,
                                Status = "A",
                                CreateDate = now,
                                LastModifiedDate = now,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            insertResults.AddRange(insertResult);
                            rowsAffected += insertResult.Count();
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = insertResults,
                            msg = $"成功新增 {rowsAffected} 筆托盤資料",
                            totalCount = rowsAffected
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

        #region //AddUserEventSetting -- 人員事件設定新增 -- Xuan 2023.07.31
        public string AddUserEventSetting(string UserEventItemName)
    {
    try
    {
        using (TransactionScope transactionScope = new TransactionScope())
        {
            using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
            {

                #region //判斷人員事件名稱是否重複
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TOP 1 1
                        FROM MES.UserEventItem
                        WHERE UserEventItemName=@UserEventItemName
                        AND CompanyId = @CompanyId";                        
                dynamicParameters.Add("UserEventItemName", UserEventItemName);
                dynamicParameters.Add("CompanyId", CurrentCompany);
                var result = sqlConnection.Query(sql, dynamicParameters);
                if (result.Count()!=0) throw new SystemException("事件資料重複!");                        
                #endregion

                #region
                dynamicParameters = new DynamicParameters();
                int existingRowCount = sqlConnection.QuerySingle<int>("SELECT COUNT(*) FROM MES.UserEventItem");
                int nextEventNumber = existingRowCount + 1;
                string str = "UserEvent";

                sql = @"INSERT INTO MES.UserEventItem (CompanyId, UserEventItemNo, UserEventItemName, Status
                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                        OUTPUT INSERTED.UserEventItemId
                        VALUES (@CompanyId, @UserEventItemNo, @UserEventItemName, @Status, @CreateDate
                        , @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                dynamicParameters.AddDynamicParams(
                new
                {
                    CompanyId = CurrentCompany,
                    UserId = CreateBy,
                    UserEventItemNo = $"{str}{nextEventNumber:000}",
                    UserEventItemName,
                    Status = "A", //啟用
                    CreateDate,
                    LastModifiedDate,
                    CreateBy,
                    LastModifiedBy = -1
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

        #region //AddMachineEventSetting -- 機台事件設定新增 -- Xuan 2023.07.31
        public string AddMachineEventSetting(int MachineId,string MachineEventNo, string MachineEventName)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷機台Id是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine
                                WHERE MachineId=@MachineId";
                        dynamicParameters.Add("MachineId", MachineId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("機台Id不存在!");
                        #endregion

                        #region //判斷機台事件名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEventItem
                                WHERE MachineEventName=@MachineEventName
                                and MachineId=@MachineId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MachineEventName", MachineEventName);
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() != 0) throw new SystemException("事件資料重複!");
                        #endregion

                        #region
                        dynamicParameters = new DynamicParameters();
                        int existingRowCount = sqlConnection.QuerySingle<int>("SELECT COUNT(*) FROM MES.MachineEventItem");
                        int nextEventNumber = existingRowCount + 1;
                        string str = "MachineEvent";
                       
                        sql = @"INSERT INTO MES.MachineEventItem (CompanyId, MachineId, MachineEventNo, MachineEventName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MachineEventItemId
                                VALUES (@CompanyId, @MachineId, @MachineEventNo, @MachineEventName, @Status, @CreateDate
                                , @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            CompanyId = CurrentCompany,
                            MachineId,
                            MachineEventNo= $"{str}{nextEventNumber:000}",
                            MachineEventName,
                            Status = "A", //啟用
                            CreateDate,
                            LastModifiedDate,
                            CreateBy = -1,
                            LastModifiedBy = -1
                        });



                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            //data = insertResult
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

        #region //AddProcessEventSetting -- 加工事件設定新增 -- Xuan 2023.08.15
        public string AddProcessEventSetting(int ParameterId, string ProcessEventNo, string ProcessEventName, string ProcessEventType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ParameterId <= 0) throw new SystemException("【製程】不能為空!");
                        if (ProcessEventName.Length <= 0) throw new SystemException("【事件名稱】不能為空!");
                        if (ProcessEventName.Length > 20) throw new SystemException("【事件名稱】長度錯誤!");
                        if (ProcessEventType.Length <= 0) throw new SystemException("【加工類別】不能為空!");

                        #region //判斷製程參數Id是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessParameter
                                WHERE ParameterId=@ParameterId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("製程參數Id不存在!");
                        #endregion

                        #region //判斷加工事件名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessEventItem
                                WHERE ProcessEventName=@ProcessEventName
                                and ParameterId=@ParameterId";
                        dynamicParameters.Add("ProcessEventName", ProcessEventName);
                        dynamicParameters.Add("ParameterId", ParameterId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() != 0) throw new SystemException("事件資料重複!");
                        #endregion

                        #region
                        dynamicParameters = new DynamicParameters();
                        int existingRowCount = sqlConnection.QuerySingle<int>("SELECT COUNT(*) FROM MES.ProcessEventItem");
                        int nextEventNumber = existingRowCount + 1;
                        string str = "ProcessEvent";

                        sql = @"INSERT INTO MES.ProcessEventItem (ParameterId, ProcessEventNo, ProcessEventName, ProcessEventType
                                ,Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProcessEventItemId
                                VALUES (@ParameterId, @ProcessEventNo, @ProcessEventName, @ProcessEventType, @Status, @CreateDate
                                , @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            ParameterId,
                            ProcessEventNo = $"{str}{nextEventNumber:000}",
                            ProcessEventName,
                            ProcessEventType,
                            Status = "A", //啟用
                            CreateDate,
                            LastModifiedDate,
                            CreateBy,
                            LastModifiedBy,
                        });



                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            //data = insertResult
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

        #region //AddDefaultRoutingQcitem 途程品號帶入製程量測項目(新增品號時) --GPai 20240513
        public string AddDefaultRoutingQcitem(int RoutingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        int routingItemId = 0;
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
                            #region //判斷途程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                    FROM MES.Routing
                                    WHERE RoutingId = @RoutingId";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("途程資料錯誤!");
                            #endregion

                            #region //找最新途程品號資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.RoutingItemId
                                    , a.RoutingId, a.ControlId, a.MtlItemId, a.Status, a.RoutingItemConfirm
                                    , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                                    , f.RoutingName, f.ModeId
                                    , g.ModeNo, g.ModeName
						            FROM MES.RoutingItem a
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN MES.Routing f ON a.RoutingId = f.RoutingId
                                    INNER JOIN MES.ProdMode g ON f.ModeId = g.ModeId
                                    WHERE a.RoutingId = @RoutingId
                                    ORDER BY a.RoutingItemId DESC";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                            foreach (var item in result2)
                            {
                                routingItemId = item.RoutingItemId;

                                #region //途程品號製程

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ItemProcessId
                                        , a.RoutingItemId, a.RoutingProcessId,a.AttrSetting
                                        , ISNULL(a.RoutingItemProcessDesc, '') RoutingItemProcessDesc
                                        , ISNULL(a.Remark, '') Remark
                                        , a.DisplayStatus, a.CycleTime, a.MoveTime
                                        , b.ProcessId, b.SortNumber, b.ProcessAlias
                                        , c.ProcessNo, c.ProcessName
                                        , d.RoutingName
                                        , f.MtlItemNo
                                        , a.ProcessingTime, a.WatingTime
                                        , g1.ModeId, g1.ParameterId
                                        FROM MES.RoutingItemProcess a
                                        INNER JOIN MES.RoutingProcess b ON a.RoutingProcessId = b.RoutingProcessId
                                        INNER JOIN MES.Process c ON b.ProcessId = c.ProcessId
                                        INNER JOIN MES.Routing d ON b.RoutingId = d.RoutingId
                                        INNER JOIN MES.RoutingItem e ON a.RoutingItemId = e.RoutingItemId
                                        INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
						                LEFT JOIN MES.ProcessParameter g1 ON g1.ProcessId = c.ProcessId
                                        WHERE c.CompanyId = @CompanyId AND a.RoutingItemId = @RoutingItemId AND g1.ModeId = @ModeId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("RoutingItemId", item.RoutingItemId);
                                dynamicParameters.Add("ModeId", item.ModeId);



                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() <= 0) throw new SystemException("途程品號製程資料錯誤!");

                                #region //找製程參數量測項目
                                foreach (var item2 in result3)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT *
                                            FROM MES.ProcessParameterQcItem
                                            WHERE ParameterId = @ParameterId";
                                    dynamicParameters.Add("ParameterId", item2.ParameterId);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result4.Count() > 0)
                                    {
                                        foreach (var item3 in result4)
                                        {
                                            #region //INSERT
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO MES.RoutingItemQcItem (RoutingItemId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, Remark
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, ItemProcessId, QcItemDesc, BallMark, Unit, QmmDetailId)
                                                    OUTPUT INSERTED.RoutingItemQcItemId
                                                    VALUES (@RoutingItemId, @QcItemId, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ItemProcessId, @QcItemDesc, @BallMark, @Unit, @QmmDetailId)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    RoutingItemId = routingItemId,
                                                    QcItemId = item3.QcItemId,
                                                    DesignValue = item3.DesignValue,
                                                    UpperTolerance = item3.UpperTolerance,
                                                    LowerTolerance = item3.LowerTolerance,
                                                    Remark = item3.Remark,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    ItemProcessId = item2.ItemProcessId,
                                                    QcItemDesc = item3.QcItemDesc,
                                                    BallMark = item3.BallMark,
                                                    Unit = item3.Unit,
                                                    QmmDetailId = item3.QmmDetailId
                                                });
                                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                            rowsAffected += insertResult.Count();
                                            #endregion

                                        }
                                    }

                                    #endregion

                                }
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

        #region //AddDefaultRoutingQcitem2 途程品號帶入製程量測項目(Excel帶入) --GPai 20240513
        public string AddDefaultRoutingQcitem2(int RoutingItemId, int ItemProcessId)//RoutingItemId ItemProcessId
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        int routingItemId = 0;
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
                            #region //找途程品號資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.RoutingItemId
                                    , a.RoutingId, a.ControlId, a.MtlItemId, a.Status, a.RoutingItemConfirm
                                    , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                                    , f.RoutingName, f.ModeId
                                    , g.ModeNo, g.ModeName
						            FROM MES.RoutingItem a
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN MES.Routing f ON a.RoutingId = f.RoutingId
                                    INNER JOIN MES.ProdMode g ON f.ModeId = g.ModeId
                                    WHERE a.RoutingItemId = @RoutingItemId";
                            dynamicParameters.Add("RoutingItemId", RoutingItemId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                            foreach (var item in result2)
                            {

                                #region //途程品號製程

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ItemProcessId
                                        , a.RoutingItemId, a.RoutingProcessId,a.AttrSetting
                                        , ISNULL(a.RoutingItemProcessDesc, '') RoutingItemProcessDesc
                                        , ISNULL(a.Remark, '') Remark
                                        , a.DisplayStatus, a.CycleTime, a.MoveTime
                                        , b.ProcessId, b.SortNumber, b.ProcessAlias
                                        , c.ProcessNo, c.ProcessName
                                        , d.RoutingName
                                        , f.MtlItemNo
                                        , a.ProcessingTime, a.WatingTime
                                        , g1.ModeId, g1.ParameterId
                                        FROM MES.RoutingItemProcess a
                                        INNER JOIN MES.RoutingProcess b ON a.RoutingProcessId = b.RoutingProcessId
                                        INNER JOIN MES.Process c ON b.ProcessId = c.ProcessId
                                        INNER JOIN MES.Routing d ON b.RoutingId = d.RoutingId
                                        INNER JOIN MES.RoutingItem e ON a.RoutingItemId = e.RoutingItemId
                                        INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
						                LEFT JOIN MES.ProcessParameter g1 ON g1.ProcessId = c.ProcessId
                                        WHERE c.CompanyId = @CompanyId AND a.RoutingItemId = @RoutingItemId AND a.ItemProcessId  = @ItemProcessId AND g1.ModeId = @ModeId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("RoutingItemId", RoutingItemId);
                                dynamicParameters.Add("ItemProcessId", ItemProcessId);
                                dynamicParameters.Add("ModeId", item.ModeId);



                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() <= 0) throw new SystemException("途程品號製程資料錯誤!");

                                #region //找製程參數量測項目
                                foreach (var item2 in result3)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT *
                                            FROM MES.ProcessParameterQcItem
                                            WHERE ParameterId = @ParameterId";
                                    dynamicParameters.Add("ParameterId", item2.ParameterId);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);
                                    
                                    if (result4.Count() > 0)
                                    {
                                        
                                        foreach (var item3 in result4)
                                        {

                                            #region // 找同量測項目
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT *
                                            FROM MES.RoutingItemQcItem
                                            WHERE QcItemId = @QcItemId AND ItemProcessId = @ItemProcessId";
                                            dynamicParameters.Add("QcItemId", item3.QcItemId);
                                            dynamicParameters.Add("ItemProcessId", ItemProcessId);
                                            var result5 = sqlConnection.Query(sql, dynamicParameters);

                                            #endregion
                                            if (result5.Count() <= 0) {
                                                #region //INSERT
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"INSERT INTO MES.RoutingItemQcItem (RoutingItemId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, Remark
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, ItemProcessId, QcItemDesc, BallMark, Unit, QmmDetailId)
                                                    OUTPUT INSERTED.RoutingItemQcItemId
                                                    VALUES (@RoutingItemId, @QcItemId, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ItemProcessId, @QcItemDesc, @BallMark, @Unit, @QmmDetailId)";
                                                dynamicParameters.AddDynamicParams(
                                                    new
                                                    {
                                                        RoutingItemId,
                                                        QcItemId = item3.QcItemId,
                                                        DesignValue = item3.DesignValue,
                                                        UpperTolerance = item3.UpperTolerance,
                                                        LowerTolerance = item3.LowerTolerance,
                                                        Remark = item3.Remark,
                                                        CreateDate,
                                                        LastModifiedDate,
                                                        CreateBy,
                                                        LastModifiedBy,
                                                        ItemProcessId,
                                                        QcItemDesc = item3.QcItemDesc,
                                                        BallMark = item3.BallMark,
                                                        Unit = item3.Unit,
                                                        QmmDetailId = item3.QmmDetailId
                                                    });
                                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                                rowsAffected += insertResult.Count();
                                                #endregion

                                            }

                                        }
                                    }

                                    #endregion

                                }
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

        #region //AddDefaultRoutingQcitem3 途程品號帶入製程量測項目(確認製程時 全部途程品號新增) --GPai 20240513
        public string AddDefaultRoutingQcitem3(int RoutingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        int routingItemId = 0;
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
                            #region //判斷途程資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                    FROM MES.Routing
                                    WHERE RoutingId = @RoutingId";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("途程資料錯誤!");
                            #endregion

                            #region //找所有途程品號資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RoutingItemId
                                    , a.RoutingId, a.ControlId, a.MtlItemId, a.Status, a.RoutingItemConfirm
                                    , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                                    , f.RoutingName, f.ModeId
                                    , g.ModeNo, g.ModeName
						            FROM MES.RoutingItem a
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN MES.Routing f ON a.RoutingId = f.RoutingId
                                    INNER JOIN MES.ProdMode g ON f.ModeId = g.ModeId
                                    WHERE a.RoutingId = @RoutingId
                                    ORDER BY a.RoutingItemId DESC";
                            dynamicParameters.Add("RoutingId", RoutingId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("途程品號資料錯誤!");
                            foreach (var item in result2)
                            {
                                routingItemId = item.RoutingItemId;

                                #region //途程品號製程

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ItemProcessId
                                        , a.RoutingItemId, a.RoutingProcessId,a.AttrSetting
                                        , ISNULL(a.RoutingItemProcessDesc, '') RoutingItemProcessDesc
                                        , ISNULL(a.Remark, '') Remark
                                        , a.DisplayStatus, a.CycleTime, a.MoveTime
                                        , b.ProcessId, b.SortNumber, b.ProcessAlias
                                        , c.ProcessNo, c.ProcessName
                                        , d.RoutingName
                                        , f.MtlItemNo
                                        , a.ProcessingTime, a.WatingTime
                                        , g1.ModeId, g1.ParameterId
                                        FROM MES.RoutingItemProcess a
                                        INNER JOIN MES.RoutingProcess b ON a.RoutingProcessId = b.RoutingProcessId
                                        INNER JOIN MES.Process c ON b.ProcessId = c.ProcessId
                                        INNER JOIN MES.Routing d ON b.RoutingId = d.RoutingId
                                        INNER JOIN MES.RoutingItem e ON a.RoutingItemId = e.RoutingItemId
                                        INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
						                LEFT JOIN MES.ProcessParameter g1 ON g1.ProcessId = c.ProcessId
                                        WHERE c.CompanyId = @CompanyId AND a.RoutingItemId = @RoutingItemId AND g1.ModeId = @ModeId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("RoutingItemId", item.RoutingItemId);
                                dynamicParameters.Add("ModeId", item.ModeId);



                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() > 0) {
                                    #region //找製程參數量測項目
                                    foreach (var item2 in result3)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT *
                                            FROM MES.ProcessParameterQcItem
                                            WHERE ParameterId = @ParameterId";
                                        dynamicParameters.Add("ParameterId", item2.ParameterId);

                                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result4.Count() > 0)
                                        {
                                            foreach (var item3 in result4)
                                            {

                                                #region // 找同量測項目
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT *
                                            FROM MES.RoutingItemQcItem
                                            WHERE QcItemId = @QcItemId AND ItemProcessId = @ItemProcessId";
                                                dynamicParameters.Add("QcItemId", item3.QcItemId);
                                                dynamicParameters.Add("ItemProcessId", item2.ItemProcessId);
                                                var result5 = sqlConnection.Query(sql, dynamicParameters);

                                                #endregion
                                                if (result5.Count() <= 0)
                                                {
                                                    #region //INSERT
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"INSERT INTO MES.RoutingItemQcItem (RoutingItemId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, Remark
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, ItemProcessId, QcItemDesc, BallMark, Unit, QmmDetailId)
                                                    OUTPUT INSERTED.RoutingItemQcItemId
                                                    VALUES (@RoutingItemId, @QcItemId, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ItemProcessId, @QcItemDesc, @BallMark, @Unit, @QmmDetailId)";
                                                    dynamicParameters.AddDynamicParams(
                                                        new
                                                        {
                                                            RoutingItemId = routingItemId,
                                                            QcItemId = item3.QcItemId,
                                                            DesignValue = item3.DesignValue,
                                                            UpperTolerance = item3.UpperTolerance,
                                                            LowerTolerance = item3.LowerTolerance,
                                                            Remark = item3.Remark,
                                                            CreateDate,
                                                            LastModifiedDate,
                                                            CreateBy,
                                                            LastModifiedBy,
                                                            ItemProcessId = item2.ItemProcessId,
                                                            QcItemDesc = item3.QcItemDesc,
                                                            BallMark = item3.BallMark,
                                                            Unit = item3.Unit,
                                                            QmmDetailId = item.QmmDetailId
                                                        });
                                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                                    rowsAffected += insertResult.Count();
                                                    #endregion
                                                }

                                            }
                                        }

                                        #endregion

                                    }


                                }
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

        #region //AddLensCarrierRing -- 套環資料新增 -- Jean 2025.07.01
        public string AddLensCarrierRing(string ModelName, string RingName, string Remarks, string HoleCount
            , string RingSpec, string RingCode, string RingShape, string Customer, decimal DailyDemand, decimal SafetyStock)
        {
            try
            {
                if (ModelName.Length <= 0) throw new SystemException("【機種名】不能為空!");
                if (RingCode.Length <= 0) throw new SystemException("【套環編碼】不能為空!");
                if (!int.TryParse(HoleCount, out int holeCountValue)) throw new SystemException("【孔數】必須為整數!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機種名是否重複
                        sql = @"SELECT TOP 1 1
                                FROM MES.LensCarrierRing
                                WHERE [Status]='A' AND  ModelName = @ModelName";
                        dynamicParameters.Add("ModelName", ModelName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機種名】重複，請重新輸入!");
                        #endregion

                        //#region //判斷套環編號是否重複
                        //sql = @"SELECT TOP 1 1
                        //        FROM MES.LensCarrierRing
                        //        WHERE  [Status]='A' AND RingCode = @RingCode";
                        //dynamicParameters.Add("RingCode", RingCode);

                        //var result2 = sqlConnection.Query(sql, dynamicParameters);
                        //if (result2.Count() > 0) throw new SystemException("【套環編號】重複，請重新輸入!");
                        //#endregion

                        #region //判斷機種名+套環編碼是否重複
                        sql = @"SELECT TOP 1 1
                        FROM MES.LensCarrierRing
                        WHERE  [Status]='A' AND ModelName = @ModelName AND RingCode = @RingCode";
                        dynamicParameters = new DynamicParameters(); // 重置参数避免污染
                        dynamicParameters.Add("ModelName", ModelName);
                        dynamicParameters.Add("RingCode", RingCode);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0)
                            throw new SystemException("【機種名+套環編碼】組合已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.LensCarrierRing (ModelName, RingName, Remarks, HoleCount, RingSpec, RingCode, RingShape, Customer, DailyDemand, SafetyStock, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RingId
                                VALUES (@ModelName, @RingName, @Remarks, @HoleCount, @RingSpec, @RingCode, @RingShape, @Customer, @DailyDemand, @SafetyStock, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModelName,
                                RingName,
                                Remarks,
                                HoleCount,
                                RingSpec,
                                RingCode,
                                RingShape,
                                Customer,
                                DailyDemand,
                                SafetyStock,
                                Status = "A", //啟用
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

        #region //AddLensCarrierRing -- 批量新增套環資料 -- Jean 2025.07.04
        public string AddLensCarrierRing(List<LCRExcelFormat> lcrExcelFormats)
        {
            try
            {

                int RingId = -1;
                int rowsAffected = 0;

                string mesErr = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        foreach (var lcrExcel in lcrExcelFormats)
                        {
                            if (string.IsNullOrWhiteSpace(lcrExcel.ModelName))
                                throw new SystemException($"【機種名】不能為空! 行數: {lcrExcel.RowNumber + 2}"); // 假设LCRExcelFormat有行号属性
                            if (string.IsNullOrWhiteSpace(lcrExcel.RingCode))
                                throw new SystemException($"【套環編碼】不能為空! 行數: {lcrExcel.RowNumber + 2}");

                            // 检查机种名是否重复
                            var modelNameCheckSql = @"SELECT TOP 1 ModelName 
                                           FROM MES.LensCarrierRing 
                                           WHERE [Status]='A' AND ModelName = @ModelName";
                            var existingModelName = sqlConnection.QueryFirstOrDefault<string>(modelNameCheckSql,
                                new { ModelName = lcrExcel.ModelName });
                            if (existingModelName != null)
                                throw new SystemException($"【機種名】重複: {lcrExcel.ModelName}");

                            //// 检查套环编号是否重复
                            //var ringCodeCheckSql = @"SELECT TOP 1 RingCode 
                            //               FROM MES.LensCarrierRing 
                            //               WHERE [Status]='A' AND RingCode = @RingCode";
                            //var existingRingCode = sqlConnection.QueryFirstOrDefault<string>(ringCodeCheckSql,
                            //    new { RingCode = lcrExcel.RingCode });
                            //if (existingRingCode != null)
                            //    throw new SystemException($"【套環編號】重複: {lcrExcel.RingCode}");

                            // 检查组合是否重复
                            var comboCheckSql = @"SELECT TOP 1 ModelName, RingCode 
                                        FROM MES.LensCarrierRing 
                                        WHERE [Status]='A' AND ModelName = @ModelName AND RingCode = @RingCode";
                            var existingCombo = sqlConnection.QueryFirstOrDefault<dynamic>(comboCheckSql,
                                new { lcrExcel.ModelName, lcrExcel.RingCode });
                            if (existingCombo != null)
                                throw new SystemException($"【機種名+套環編碼】組合已存在: {lcrExcel.ModelName} + {lcrExcel.RingCode}");


                            #region //INSERT MES.LensCarrierRing

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.LensCarrierRing (ModelName, RingName, Remarks, HoleCount, RingSpec, RingCode, RingShape, Customer, DailyDemand, SafetyStock, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RingId
                                    VALUES (@ModelName, @RingName, @Remarks, @HoleCount, @RingSpec, @RingCode, @RingShape, @Customer, @DailyDemand, @SafetyStock, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)
                                    ";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                RingId,
                                lcrExcel.ModelName,
                                lcrExcel.RingName,
                                lcrExcel.Remarks,
                                lcrExcel.HoleCount,
                                lcrExcel.RingSpec,
                                lcrExcel.RingCode,
                                lcrExcel.RingShape,
                                lcrExcel.Customer,
                                lcrExcel.DailyDemand,
                                lcrExcel.SafetyStock,
                                Status = "A",
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
                            err = mesErr
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

        #region //AddLensCarrierRing -- 套環庫存異動新增 -- Jean 2025.07.07
        public string AddLensCarrierRing(int RingId, string ModelName, string RingName, string TransType
            , string TransDate, int Quantity)
        {
            try
            {
                if (RingId <= 0) throw new SystemException("【套環編碼】不能為空!");
                if (ModelName.Length <= 0) throw new SystemException("【機種名】不能為空!");
                if (TransType.Length <= 0) throw new SystemException("【異動類型】不能為空!");
                if (TransDate.Length <= 0) throw new SystemException("【異動日期】不能為空!");
                if (Quantity <= 0) throw new SystemException("【數量】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷新增时套环库存是否出现负库存情况
                        sql = @"SELECT x.Quantity
                                FROM (
                                    SELECT a.RingId,ISNULL(SUM(c.Quantity),0)Quantity
                                    FROM  MES.LensCarrierRing a
                                    LEFT JOIN (SELECT e.RingId,(CASE WHEN e.TransType='IN' THEN e.Quantity WHEN e.TransType='OUT' THEN -e.Quantity ELSE 0 END) Quantity
			                                    FROM MES.RingTransaction e
			                                    WHERE e.Status='A'
		                                    )c ON c.RingId = a.RingId
                                    INNER JOIN BAS.[User] d on a.CreateBy = d.UserId
                                    WHERE 1=1
                                    GROUP BY a.RingId
                                    )x
                                WHERE x.RingId=@RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var currentStock = sqlConnection.QuerySingleOrDefault<int?>(sql, dynamicParameters) ?? 0;

                        int newStock = TransType == "IN"
                            ? currentStock + Quantity
                            : currentStock - Quantity;

                        // 检查是否会导致负库存
                        if (newStock < 0)
                        {
                            throw new SystemException($"當前庫存: {currentStock}，庫存不足，此操作會導致負庫存! ");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.RingTransaction (RingId, ModelName, RingName, TransType, TransDate, Quantity, Status
                                 , CreateDate,  CreateBy)
                                OUTPUT INSERTED.RingTransId
                                VALUES (@RingId, @ModelName, @RingName, @TransType, @TransDate, @Quantity, @Status
                                , @CreateDate, @CreateBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RingId,
                                ModelName,
                                RingName,
                                TransType,
                                TransDate,
                                Quantity,
                                Status = "A", //啟用
                                CreateDate,
                                CreateBy
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

        #region //AddRingTransaction -- 批量新增套環庫存異動資料 -- Jean 2025.07.08
        public string AddRingTransaction(List<LCRTransExcelFormat> lcrTransExcelFormats)
        {
            try
            {
                int rowsAffected = 0;
                string mesErr = "";

                // 第一步：预处理Excel数据，计算每个套环的总库存变化
                // 存储结构：Key=RingId，Value=该套环在Excel中的总库存变化量
                var excelRingStockChanges = new Dictionary<int, int>();
                // 存储RingId与ModelName的映射，用于错误提示
                var ringIdToModelName = new Dictionary<int, string>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlConnection.Open();

                    // 预处理Excel数据，验证基础信息并计算库存变化
                    foreach (var lcrTransExcel in lcrTransExcelFormats)
                    {
                        // 基础校验
                        if (string.IsNullOrWhiteSpace(lcrTransExcel.ModelName))
                            throw new SystemException($"【機種名】不能為空! 行數: {lcrTransExcel.RowNumber + 2}");
                        if (string.IsNullOrWhiteSpace(lcrTransExcel.TransType))
                            throw new SystemException($"【異動類型】不能為空! 行數: {lcrTransExcel.RowNumber + 2}");
                        if (string.IsNullOrWhiteSpace(lcrTransExcel.TransDate))
                            throw new SystemException($"【異動日期】不能為空! 行數: {lcrTransExcel.RowNumber + 2}");
                        if (lcrTransExcel.Quantity <= 0)
                            throw new SystemException($"【數量】必須大於0! 行數: {lcrTransExcel.RowNumber + 2}");

                        // 通过ModelName查询RingId
                        string selectSql = @"SELECT RingId,RingName FROM MES.LensCarrierRing WHERE ModelName = @ModelName AND Status='A' ";
                        var dynamicParameters = new DynamicParameters();
                        dynamicParameters.Add("@ModelName", lcrTransExcel.ModelName);

                        var ringResult = sqlConnection.Query<dynamic>(selectSql, dynamicParameters).FirstOrDefault();

                        if (ringResult == null)
                        {
                            throw new SystemException($"找不到匹配的套環資料，套環編號: {lcrTransExcel.ModelName}, 行數: {lcrTransExcel.RowNumber + 2}");
                        }

                        int ringId = ringResult.RingId;
                        string ringName = ringResult.RingName;

                        // 记录RingId与ModelName的映射
                        if (!ringIdToModelName.ContainsKey(ringId))
                        {
                            ringIdToModelName[ringId] = lcrTransExcel.ModelName;
                        }

                        // 计算当前行对库存的影响
                        int currentRowEffect = lcrTransExcel.TransType == "IN"
                            ? lcrTransExcel.Quantity
                            : -lcrTransExcel.Quantity;

                        // 累加该套环在Excel中的总库存变化
                        if (excelRingStockChanges.ContainsKey(ringId))
                        {
                            excelRingStockChanges[ringId] += currentRowEffect;
                        }
                        else
                        {
                            excelRingStockChanges[ringId] = currentRowEffect;
                        }
                    }

                    // 第二步：与系统库存比对，检查是否会出现负库存
                    foreach (var ringStock in excelRingStockChanges)
                    {
                        int ringId = ringStock.Key;
                        int totalExcelEffect = ringStock.Value;
                        string modelName = ringIdToModelName[ringId];

                        // 查询系统当前库存
                        string stockSql = @"SELECT x.Quantity
                                    FROM (
                                        SELECT a.RingId,ISNULL(SUM(c.Quantity),0) AS Quantity
                                        FROM  MES.LensCarrierRing a
                                        LEFT JOIN (SELECT e.RingId,(CASE WHEN e.TransType='IN' THEN e.Quantity WHEN e.TransType='OUT' THEN -e.Quantity ELSE 0 END) Quantity
                                                    FROM MES.RingTransaction e
                                                    WHERE e.Status='A'
                                                )c ON c.RingId = a.RingId
                                        INNER JOIN BAS.[User] d on a.CreateBy = d.UserId
                                        WHERE a.RingId = @RingId
                                        GROUP BY a.RingId
                                    )x";

                        var dynamicParameters = new DynamicParameters();
                        dynamicParameters.Add("RingId", ringId);

                        var currentSystemStock = sqlConnection.QuerySingleOrDefault<int?>(stockSql, dynamicParameters) ?? 0;

                        // 计算处理后预计库存
                        int projectedStock = currentSystemStock + totalExcelEffect;

                        // 检查负库存
                        if (projectedStock < 0)
                        {
                            throw new SystemException($"批量處理會導致負庫存! 套環編號: {modelName}, 系統當前庫存: {currentSystemStock}, 批量處理後預計庫存: {projectedStock}");
                        }
                    }

                    // 第三步：所有校验通过，执行批量插入
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        foreach (var lcrTransExcel in lcrTransExcelFormats)
                        {
                            // 再次查询RingId（确保数据一致性，也可优化为缓存）
                            string selectSql = @"SELECT RingId,RingName FROM MES.LensCarrierRing WHERE ModelName = @ModelName AND Status='A' ";
                            var dynamicParameters = new DynamicParameters();
                            dynamicParameters.Add("@ModelName", lcrTransExcel.ModelName);

                            var ringResult = sqlConnection.Query<dynamic>(selectSql, dynamicParameters).FirstOrDefault();
                            int ringId = ringResult.RingId;
                            string ringName = ringResult.RingName;

                            // 插入记录
                            dynamicParameters = new DynamicParameters();
                            string sql = @"INSERT INTO MES.RingTransaction (RingId, ModelName, RingName, TransType, TransDate, Quantity, Status
                                    , CreateDate, CreateBy)
                                    OUTPUT INSERTED.RingTransId
                                    VALUES (@RingId, @ModelName, @RingName, @TransType, @TransDate, @Quantity, @Status
                                    , @CreateDate, @CreateBy)
                                    ";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                RingId = ringId,
                                lcrTransExcel.ModelName,
                                RingName = ringName,
                                lcrTransExcel.TransType,
                                TransDate = DateTime.Parse(lcrTransExcel.TransDate),
                                lcrTransExcel.Quantity,
                                Status = "A",
                                CreateDate,
                                CreateBy
                            });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();
                        }

                        transactionScope.Complete();
                    }

                    // 响应结果
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = $"({rowsAffected} rows affected)",
                        err = mesErr
                    });
                }
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #endregion

        #region //Update
        #region //UpdateProdMode -- 生產模式更新 -- Zoey 2022.05.20
        public string UpdateProdMode(int ModeId, string ModeName, string ModeDesc, string BarcodeCtrl , string ScrapRegister
            , string Source, string PVTQCFlag, string NgToBarcode, string TrayBarcode, string LotStatus, string OutputBarcodeFlag, string MrType, string OQcCheckType
            )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ModeName.Length <= 0) throw new SystemException("【模式名稱】不能為空!");
                        if (ModeName.Length > 100) throw new SystemException("【模式名稱】長度錯誤!");
                        if (ModeDesc.Length <= 0) throw new SystemException("【模式描述】不能為空!");
                        if (ModeDesc.Length > 100) throw new SystemException("【模式描述】長度錯誤!");
                        if (BarcodeCtrl.Length <= 0) throw new SystemException("【過站是否要條碼控制】不能為空!");
                        if (ScrapRegister.Length <= 0) throw new SystemException("【報廢品是否入報廢倉】不能為空!");

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProdMode SET
                                ModeName = @ModeName,
                                ModeDesc = @ModeDesc,
                                BarcodeCtrl = @BarcodeCtrl,
                                ScrapRegister = @ScrapRegister,
                                Source = @Source,
                                PVTQCFlag = @PVTQCFlag,
                                NgToBarcode = @NgToBarcode,
                                TrayBarcode = @TrayBarcode,
                                LotStatus = @LotStatus,
                                OutputBarcodeFlag = @OutputBarcodeFlag,
                                MrType = @MrType,
                                OQcCheckType = @OQcCheckType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModeId = @ModeId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModeName,
                                ModeDesc,
                                BarcodeCtrl,
                                ScrapRegister,
                                Source,
                                PVTQCFlag,
                                NgToBarcode,
                                TrayBarcode,
                                LotStatus,
                                OutputBarcodeFlag,
                                MrType,
                                OQcCheckType,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModeId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProdModeStatus -- 生產模式狀態更新 -- Zoey 2022.05.20
        public string UpdateProdModeStatus(int ModeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷生產模式資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ProdMode
                                WHERE ModeId = @ModeId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProdMode SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModeId = @ModeId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModeId,
                                CompanyId = CurrentCompany,
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProdModeShift -- 生產模式班次資料更新 -- Shintokru 2022.06.20
        public string UpdateProdModeShift(int ModeShiftId, int ShiftId, string EffectiveDate, string ExpirationDate)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ShiftId < 0) throw new SystemException("【班次】不能為空!");
                        if (EffectiveDate.Length <= 0) throw new SystemException("【生效日】不能為空!");
                        if (ExpirationDate.Length > 0)
                        {
                            if (EffectiveDate.Equals(ExpirationDate)) throw new SystemException("【生效日】與【失效日】不能相同!");
                            if (DateTime.Compare(Convert.ToDateTime(EffectiveDate), Convert.ToDateTime(ExpirationDate)) > 0) throw new SystemException("【生效日】不能大於【失效日】!");
                        }

                        string nullData = null;

                        #region //判斷生產模式班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdModeShift
                                WHERE ModeShiftId = @ModeShiftId";
                        dynamicParameters.Add("ModeShiftId", ModeShiftId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式班次資料錯誤!");
                        #endregion

                        #region //判斷班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Shift]
                                WHERE CompanyId = @CompanyId
                                AND ShiftId = @ShiftId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShiftId", ShiftId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("班次資料錯誤!");
                        #endregion

                        #region //判斷生產模式班次資料是否重複(因沒有傳遞ModeId,所以寫子查詢把ModeId撈出來)
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdModeShift
                                WHERE ModeId !=
                                    (
                                        SELECT ModeId
                                        FROM MES.ProdModeShift
                                        WHERE ModeShiftId = @ModeShiftId
                                    )
                                AND ModeShiftId = @ModeShiftId
                                AND ShiftId = @ShiftId";
                        dynamicParameters.Add("ModeShiftId", ModeShiftId);
                        dynamicParameters.Add("ShiftId", ShiftId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【生產模式班次】重複，請重新輸入!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProdModeShift SET
                                ShiftId = @ShiftId,
                                EffectiveDate = @EffectiveDate,
                                ExpirationDate = @ExpirationDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModeShiftId = @ModeShiftId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShiftId,
                                EffectiveDate,
                                ExpirationDate = ExpirationDate.Length > 0 ? ExpirationDate : nullData,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModeShiftId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProdModeShiftStatus -- 生產模式班次狀態更新 -- Shintokru 2022.06.23
        public string UpdateProdModeShiftStatus(int ModeShiftId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷生產模式班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ProdModeShift
                                WHERE ModeShiftId = @ModeShiftId";
                        dynamicParameters.Add("ModeShiftId", ModeShiftId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式班次資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProdModeShift SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModeShiftId = @ModeShiftId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModeShiftId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateWorkShop -- 車間資料更新 -- Shintokru 2022.06.07
        public string UpdateWorkShop(int ShopId, string ShopName, string ShopDesc, string Location, string Floor)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ShopName.Length <= 0) throw new SystemException("【車間名稱】不能為空!");
                        if (ShopName.Length > 100) throw new SystemException("【車間名稱】長度錯誤!");
                        if (ShopDesc.Length <= 0) throw new SystemException("【車間描述】不能為空!");
                        if (ShopDesc.Length > 100) throw new SystemException("【車間描述】長度錯誤!");
                        if (Location.Length > 100) throw new SystemException("【位置】長度錯誤!");
                        if (Floor.Length > 100) throw new SystemException("【樓層】長度錯誤!");

                        #region //判斷車間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopId = @ShopId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopId", ShopId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");
                        #endregion

                        #region //判斷車間名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopName = @ShopName
                                AND ShopId != @ShopId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopName", ShopName);
                        dynamicParameters.Add("ShopId", ShopId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【車間名稱】重複，請重新輸入!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.WorkShop SET
                                ShopName = @ShopName,
                                ShopDesc = @ShopDesc,
                                Location = @Location,
                                Floor = @Floor,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ShopId = @ShopId
                                AND CompanyId = CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShopName,
                                ShopDesc,
                                Location = Location.Length > 0 ? Location : null,
                                Floor = Floor.Length > 0 ? Floor : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ShopId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateWorkShopStatus -- 車間狀態更新 -- Shintokru 2022.06.07
        public string UpdateWorkShopStatus(int ShopId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷車間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopId = @ShopId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopId", ShopId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.WorkShop SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ShopId = @ShopId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ShopId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateMachine -- 機台資料更新 -- Ted 2022.06.07
        public string UpdateMachine(int MachineId, int ShopId, string MachineName, string MachineDesc)
        {
            try
            {
                if (ShopId <= 0) throw new SystemException("【所屬車間】不能為空!");
                if (MachineName.Length <= 0) throw new SystemException("【機台名稱】不能為空!");
                if (MachineName.Length > 100) throw new SystemException("【機台名稱】長度錯誤!");
                if (MachineDesc.Length <= 0) throw new SystemException("【機台描述】不能為空!");
                if (MachineDesc.Length > 100) throw new SystemException("【機台描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷車間資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.WorkShop
                                WHERE CompanyId = @CompanyId
                                AND ShopId = @ShopId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShopId", ShopId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");
                        #endregion

                        #region //判斷機台資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine
                                WHERE MachineId = @MachineId";
                        dynamicParameters.Add("MachineId", MachineId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Machine SET
                                ShopId = @ShopId,
                                MachineName = @MachineName,
                                MachineDesc = @MachineDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MachineId = @MachineId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShopId,
                                MachineName,
                                MachineDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                MachineId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateMachineStatus -- 機台狀態更新 -- Ted 2022.06.07
        public string UpdateMachineStatus(int MachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台資訊是否正確
                        sql = @"SELECT TOP 1 a.Status
                                FROM MES.Machine a
                                INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId
                                WHERE a.MachineId = @MachineId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Machine SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MachineId = @MachineId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                MachineId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateTimeInterval -- 時間區段更新 -- Shintoku 2022.06.15
        public string UpdateTimeInterval(int IntervalId, string BeginTime, string EndTime)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (IntervalId <= 0) throw new SystemException("【時間區段ID】不能為空!");
                        if (BeginTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                        if (EndTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");

                        #region //判斷時間區段是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.TimeInterval
                                WHERE IntervalId = @IntervalId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("IntervalId", IntervalId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("時間區段資料錯誤!");
                        #endregion

                        #region //判斷開始時間是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.TimeInterval
                                WHERE BeginTime = @BeginTime
                                AND IntervalId != @IntervalId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("BeginTime", BeginTime);
                        dynamicParameters.Add("IntervalId", IntervalId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【開始時間】重複，請重新輸入!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.TimeInterval SET
                                BeginTime = @BeginTime,
                                EndTime = @EndTime,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE IntervalId = @IntervalId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                BeginTime,
                                EndTime,
                                LastModifiedDate,
                                LastModifiedBy,
                                IntervalId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateDevice -- 裝置資料更新 -- Ted 2022.06.15
        public string UpdateDevice(int DeviceId, string DeviceType, string DeviceName , string DeviceDesc, string DeviceIdentifierCode, string DeviceAuthority
            )
        {
            try
            {
                if (DeviceType.Length <= 0) throw new SystemException("【設備類別】不能為空!");
                if (DeviceName.Length <= 0) throw new SystemException("【設備識別碼】不能為空!");
                if (DeviceName.Length > 100) throw new SystemException("【設備識別碼】長度錯誤!");
                if (DeviceDesc.Length <= 0) throw new SystemException("【設備識別碼】不能為空!");
                if (DeviceDesc.Length > 100) throw new SystemException("【設備識別碼】長度錯誤!");
                if (DeviceIdentifierCode.Length <= 0) throw new SystemException("【設備識別碼】不能為空!");
                if (DeviceIdentifierCode.Length > 50) throw new SystemException("【設備識別碼】長度錯誤!");
                if (DeviceAuthority.Length <= 0) throw new SystemException("【設備權限】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Device
                                WHERE DeviceId = @DeviceId";
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("裝置資料錯誤!");
                        #endregion

                        #region //判斷設備名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceName = @DeviceName
                                AND DeviceId != @DeviceId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceName", DeviceName);
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【設備名稱】重複，請重新輸入!");
                        #endregion

                        #region //判斷設備識別碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceIdentifierCode = @DeviceIdentifierCode
                                AND DeviceId != @DeviceId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceIdentifierCode", DeviceIdentifierCode);
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【設備識別碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Device SET
                                DeviceType = @DeviceType,
                                DeviceName = @DeviceName,
                                DeviceDesc = @DeviceDesc,
                                DeviceIdentifierCode = @DeviceIdentifierCode,
                                DeviceAuthority = @DeviceAuthority,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DeviceId = @DeviceId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DeviceType,
                                DeviceName,
                                DeviceDesc,
                                DeviceIdentifierCode,
                                DeviceAuthority,
                                LastModifiedDate,
                                LastModifiedBy,
                                DeviceId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateDeviceIdStatus -- 裝置狀態更新 -- Ted 2022.06.15
        public string UpdateDeviceIdStatus(int DeviceId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceId = @DeviceId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("裝置資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Device SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DeviceId = @DeviceId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DeviceId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateDeviceMachine -- 裝置機台綁定資料更新 -- Ted 2022.06.15
        public string UpdateDeviceMachine(int DeviceId, int MachineId
            )
        {
            try
            {
                if (DeviceId <= 0) throw new SystemException("【裝置ID】不能為空!");
                if (MachineId <= 0) throw new SystemException("【機台ID】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.DeviceMachine
                                WHERE 1=1
                                AND DeviceId = @DeviceId";
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("裝置資料錯誤!");
                        #endregion

                        #region //判斷機台是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.DeviceMachine
                                WHERE 1=1
                                AND MachineId != @MachineId";
                        dynamicParameters.Add("MachineId", MachineId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【機台】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.DeviceMachine SET
                                MachineId = @MachineId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DeviceId = @DeviceId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MachineId,
                                LastModifiedDate,
                                LastModifiedBy,
                                DeviceId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcess -- 製程資料更新 -- Shintokru 2022.06.20
        public string UpdateProcess(int ProcessId, string ProcessName, string ProcessDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ProcessName.Length <= 0) throw new SystemException("【製程名稱】不能為空!");
                        if (ProcessName.Length > 100) throw new SystemException("【製程名稱】長度錯誤!");
                        if (ProcessDesc.Length <= 0) throw new SystemException("【製程描述】不能為空!");
                        if (ProcessDesc.Length > 100) throw new SystemException("【製程描述】長度錯誤!");

                        #region //判斷製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE CompanyId = @CompanyId
                                AND ProcessId = @ProcessId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ProcessId", ProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程資料錯誤!");
                        #endregion

                        #region //判斷製程名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE CompanyId = @CompanyId
                                AND ProcessName = @ProcessName
                                AND ProcessId != @ProcessId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ProcessName", ProcessName);
                        dynamicParameters.Add("ProcessId", ProcessId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【製程名稱】重複，請重新輸入!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Process SET
                                ProcessName = @ProcessName,
                                ProcessDesc = @ProcessDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProcessId = @ProcessId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessName,
                                ProcessDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                ProcessId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcessStatus -- 製程狀態更新 -- Shintokru 2022.06.15
        public string UpdateProcessStatus(int ProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.Process
                                WHERE CompanyId = @CompanyId
                                AND ProcessId = @ProcessId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ProcessId", ProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Process SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProcessId = @ProcessId
                                AND CompanyId = CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ProcessId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProdUnit -- 生產單元資料更新 -- Shintokru 2022.06.23
        public string UpdateProdUnit(int UnitId, string UnitName, string UnitDesc, string CheckStatus)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (UnitId < 0) throw new SystemException("【單元ID】不能為空!");
                        if (UnitName.Length <= 0) throw new SystemException("【單元名稱】不能為空!");
                        if (UnitName.Length > 100) throw new SystemException("【單元名稱】長度錯誤!");
                        if (UnitDesc.Length <= 0) throw new SystemException("【單元描述】不能為空!");
                        if (UnitDesc.Length > 100) throw new SystemException("【單元描述】長度錯誤!");
                        if (CheckStatus.Length < 0) throw new SystemException("【檢核狀態】長度錯誤!");

                        #region //判斷生產單元資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdUnit
                                WHERE UnitId = @UnitId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("UnitId", UnitId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");
                        #endregion

                        #region //判斷單元名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdUnit
                                WHERE UnitName = @UnitName
                                AND UnitId != @UnitId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("UnitName", UnitName);
                        dynamicParameters.Add("UnitId", UnitId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【單元名稱】重複，請重新輸入!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProdUnit SET
                                UnitName = @UnitName,
                                UnitDesc = @UnitDesc,
                                CheckStatus = @CheckStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UnitId = @UnitId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UnitName,
                                UnitDesc,
                                CheckStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                UnitId,
                                CompanyId = CurrentCompany
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProdUnitStatus -- 生產單元狀態更新 -- Shintokru 2022.06.23
        public string UpdateProdUnitStatus(int UnitId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷生產單元是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ProdUnit
                                WHERE UnitId = @UnitId
                                AND CompanyId = CompanyId";
                        dynamicParameters.Add("UnitId", UnitId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProdUnit SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UnitId = @UnitId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                UnitId,
                                CompanyId = CurrentCompany,
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateWarehouse -- 庫房基本資料更新 -- Ted 2022.07.01
        public string UpdateWarehouse(int WarehouseId, int ShopId, string WarehouseName, string WarehouseDesc)
        {
            try
            {
                if (WarehouseName.Length <= 0) throw new SystemException("【庫房名稱】不能為空!");
                if (WarehouseName.Length > 100) throw new SystemException("【庫房名稱】長度錯誤!");
                if (WarehouseDesc.Length <= 0) throw new SystemException("【庫房描述】不能為空!");
                if (WarehouseDesc.Length > 100) throw new SystemException("【庫房描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷庫房資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Warehouse a
                                INNER JOIN MES.WorkShop b on a.ShopId = b.ShopId
                                WHERE a.WarehouseId = @WarehouseId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("WarehouseId", WarehouseId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【庫房ID】錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Warehouse SET
                                ShopId = @ShopId,
                                WarehouseName = @WarehouseName,
                                WarehouseDesc = @WarehouseDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE WarehouseId = @WarehouseId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShopId,
                                WarehouseName,
                                WarehouseDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                WarehouseId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateWarehouseStatus -- 庫房基本資料狀態更新 -- Ted 2022.07.01
        public string UpdateWarehouseStatus(int WarehouseId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產模式資訊是否正確
                        sql = @"SELECT TOP 1 a.Status
                                FROM MES.Warehouse a
                                INNER JOIN MES.WorkShop b on a.ShopId = b.ShopId
                                WHERE a.WarehouseId = @WarehouseId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("WarehouseId", WarehouseId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Warehouse SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE WarehouseId = @WarehouseId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                WarehouseId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateWarehouseLocation -- 庫房儲位基本資料更新 -- Ted 2022.07.04
        public string UpdateWarehouseLocation(int WarehouseId, int LocationId, string LocationName, string LocationDesc)
        {
            try
            {
                if (LocationName.Length <= 0) throw new SystemException("【儲位名稱】不能為空!");
                if (LocationName.Length > 100) throw new SystemException("【儲位名稱】長度錯誤!");
                if (LocationDesc.Length <= 0) throw new SystemException("【儲位描述】不能為空!");
                if (LocationDesc.Length > 100) throw new SystemException("【儲位描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷儲位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.WarehouseLocation a
                                INNER JOIN MES.Warehouse b on a.WarehouseId = b.WarehouseId
                                INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                                WHERE a.LocationId = @LocationId
                                AND b.WarehouseId = @WarehouseId
                                AND c.CompanyId = @CompanyId";
                        dynamicParameters.Add("LocationId", LocationId);
                        dynamicParameters.Add("WarehouseId", WarehouseId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【儲位ID】錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.WarehouseLocation SET
                                WarehouseId = @WarehouseId,
                                LocationName = @LocationName,
                                LocationDesc = @LocationDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE LocationId = @LocationId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                WarehouseId,
                                LocationName,
                                LocationDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                LocationId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters); ;

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

        #region //UpdateWarehouseLocationStatus -- 庫房儲位基本資料狀態更新 -- Ted 2022.07.04
        public string UpdateWarehouseLocationStatus(int LocationId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產模式資訊是否正確
                        sql = @"SELECT TOP 1 a.Status
                                FROM MES.WarehouseLocation a
                                INNER JOIN MES.Warehouse b on a.WarehouseId = b.WarehouseId
                                INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                                WHERE a.LocationId = @LocationId
                                AND c.CompanyId = @CompanyId";
                        dynamicParameters.Add("LocationId", LocationId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.WarehouseLocation SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE LocationId = @LocationId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                LocationId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcessParameter -- 製程參數資料更新 -- Ted 2022.07.05
        public string UpdateProcessParameter(int ParameterId, int ProcessId, int ModeId, int DepartmentId, string ProcessCheckStatus, string PreCollectionStatus
            , string PostCollectionStatus, string NgToBarcode, string PassingMode, string ProcessCheckType, string ConsumeFlag)
        {
            try
            {
                if (ParameterId <= 0) throw new SystemException("【製程參數ID】不能為空!");
                if (ProcessId <= 0) throw new SystemException("【製程】不能為空!");
                if (ModeId <= 0) throw new SystemException("【生產模式】不能為空!");
                if (ProcessCheckStatus.Length > 0)
                {
                    if (ProcessCheckStatus == "Y")
                    {
                        if (ProcessCheckType.Length <= 0) throw new SystemException("【工程檢頻率】不能為空!");
                    }
                }
                if (PreCollectionStatus.Length <= 0) throw new SystemException("【是否收集開工前資訊】不能為空!");
                if (PostCollectionStatus.Length <= 0) throw new SystemException("【是否收集完工後資訊】不能為空!");
                if (NgToBarcode.Length <= 0) throw new SystemException("【是否刷取過站條碼】不能為空!");
                if (PassingMode.Length <= 0) throw new SystemException("【過站模式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷該製程下生產模式是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessParameter
                                WHERE ModeId = @ModeId
                                AND ProcessId = @ProcessId
                                AND DepartmentId = @DepartmentId
                                AND ParameterId != @ParameterId";
                        dynamicParameters.Add("ModeId", ModeId);
                        dynamicParameters.Add("ProcessId", ProcessId);
                        dynamicParameters.Add("DepartmentId", ParameterId);
                        dynamicParameters.Add("ParameterId", ParameterId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【該製程下生產模式】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessParameter SET
                                ProcessId = @ProcessId,
                                ModeId = @ModeId,
                                DepartmentId = @DepartmentId,
                                ProcessCheckStatus = @ProcessCheckStatus,
                                PreCollectionStatus = @PreCollectionStatus,
                                PostCollectionStatus = @PostCollectionStatus,
                                NgToBarcode = @NgToBarcode,
                                PassingMode = @PassingMode,
                                ProcessCheckType = @ProcessCheckType,
ConsumeFlag = @ConsumeFlag,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessId,
                                ModeId,
                                DepartmentId,
                                ProcessCheckStatus,
                                PreCollectionStatus,
                                PostCollectionStatus,
                                NgToBarcode,
                                PassingMode,
                                ProcessCheckType,
                                ConsumeFlag,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcessParameterStatus -- 製程參數資料狀態更新 -- Ted 2022.07.05
        public string UpdateProcessParameterStatus(int ParameterId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產單元是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ProcessParameter
                                WHERE ParameterId = @ParameterId";
                        dynamicParameters.Add("ParameterId", ParameterId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessParameter SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateBatchStatus -- 製程機台資料是否支援編成 -- Ted 2022.07.07
        public string UpdateBatchStatus(int ParameterId, int MachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產單元是否正確
                        sql = @"SELECT TOP 1 BatchStatus
                                FROM MES.ProcessMachine
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");

                        string BatchStatus = "";
                        foreach (var item in result)
                        {
                            BatchStatus = item.BatchStatus;
                        }

                        #region //調整為相反狀態
                        switch (BatchStatus)
                        {
                            case "Y":
                                BatchStatus = "N";
                                break;
                            case "N":
                                BatchStatus = "Y";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessMachine SET
                                BatchStatus = @BatchStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                BatchStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId,
                                MachineId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateParameterMachineStatus -- 製程機台資料狀態更新 -- Shintokuro 2024.06.04
        public string UpdateParameterMachineStatus(int ParameterId, int MachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產單元是否正確
                        sql = @"SELECT TOP 1 [Status]
                                FROM MES.ProcessMachine
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");

                        string Status = "";
                        foreach (var item in result)
                        {
                            Status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (Status)
                        {
                            case "A":
                                Status = "S";
                                break;
                            case "S":
                                Status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessMachine SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId,
                                MachineId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcessMachine -- 更新製程機台資料 -- Shintokru 2022.07.06
        public string UpdateProcessMachine(int ParameterId, int MachineId, int ToolCount, string KeyenceFlag, int KeyenceId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ParameterId < 0) throw new SystemException("【單元ID】不能為空!");
                        if (MachineId < 0) throw new SystemException("【機台ID】不能為空!");
                        if (ToolCount <= 0) throw new SystemException("【工具上限數】至少為1!");
                        if (KeyenceFlag == "Y" && KeyenceId <= 0) throw new SystemException("Keyence設備ID不能為空!");

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessMachine SET
                                ToolCount = @ToolCount,
                                KeyenceFlag = @KeyenceFlag,
                                KeyenceId = @KeyenceId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ToolCount,
                                KeyenceFlag,
                                KeyenceId = KeyenceId > 0 ? KeyenceId : (int?)null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId,
                                MachineId,
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcessProductionUnit -- 製程生產單元資料更新 -- Ted 2022.07.06
        public string UpdateProcessProductionUnit(int ParameterId, int UnitId, int SortNumber)
        {
            try
            {
                if (ParameterId < 0) throw new SystemException("【單元ID】不能為空!");
                if (UnitId < 0) throw new SystemException("【單元ID】不能為空!");
                if (SortNumber < 0) throw new SystemException("【單元ID】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessProductionUnit SET
                                SortNumber = @SortNumber,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId
                                AND UnitId = @UnitId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SortNumber,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId,
                                UnitId,
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateProcessProductionUnitStatus -- 製程生產單元狀態更新 -- Shintokru 2022.07.06
        public string UpdateProcessProductionUnitStatus(int ParameterId, int UnitId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產單元是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.ProcessProductionUnit
                                WHERE ParameterId = @ParameterId
                                AND UnitId = @UnitId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("UnitId", UnitId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessProductionUnit SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId
                                AND UnitId = @UnitId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ParameterId,
                                UnitId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateTray -- 托盤資料更新 -- Ted 2022.06.15
        public string UpdateTray(int TrayId, string TrayName, int TrayCapacity, string Remark
            )
        {
            try
            {
                if (TrayName.Length <= 0) throw new SystemException("【托盤名稱】不能為空!");
                if (TrayName.Length > 100) throw new SystemException("【托盤名稱】長度錯誤!");
                if (TrayCapacity < 0) throw new SystemException("【托盤容量】不能為負數!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷托盤資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tray
                                WHERE TrayId = @TrayId";
                        dynamicParameters.Add("TrayId", TrayId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【托盤管理】 - 托盤資料錯誤!");
                        #endregion

                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Tray SET
                                TrayName = @TrayName,
                                Remark = @Remark,
                                TrayCapacity = @TrayCapacity,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayId = @TrayId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TrayName,
                                Remark,
                                TrayCapacity,
                                LastModifiedDate,
                                LastModifiedBy,
                                TrayId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateTrayStatus -- 托盤狀態更新 -- Shintokuro 2023.02.04
        public string UpdateTrayStatus(int TrayId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷托盤資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.Tray
                                WHERE CompanyId = @CompanyId
                                AND TrayId = @TrayId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("TrayId", TrayId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【托盤管理】 - 托盤資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Tray SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayId = @TrayId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                TrayId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateLaserEngravedRingCode -- 自動套環雷刻機 托盤狀態更新 -- Luca 2024.12.04
        public string UpdateLaserEngravedRingCode(string Company,string UserNo,string TrayNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

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

                        #region //取得User資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo AND UserId !=216" ;
                        dynamicParameters.Add("UserNo", UserNo);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!!");

                        int UserId = -1;
                        foreach (var item in UserResult)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        #region//查詢 MES.TrayTemp
                        int rowsAffected = 0;
                        int? nullData = null;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId,a.TrayTempId,a.TrayTempNo,a.TrayTempName,a.TrayTempCapacity
                                  FROM MES.TrayTemp a                                     
                                 WHERE a.TrayTempNo = @TrayTempNo
                                   AND a.[Status] = 'A'
                                   AND a.CompanyId = @CompanyId";

                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("TrayTempNo", TrayNo);
                        var TrayTempResult = sqlConnection.Query(sql, dynamicParameters);
                        if (TrayTempResult.Count() > 0) {
                            foreach (var item in TrayTempResult)
                            {                               

                                #region //TrayTemp資料新增
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.Tray (CompanyId, BarcodeNo, TrayNo, TrayName, TrayCapacity, Remark, UseTimes, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)                                
                                VALUES (@CompanyId, @BarcodeNo, @TrayNo, @TrayName, @TrayCapacity, @Remark, @UseTimes, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.CompanyId,
                                        BarcodeNo= nullData,
                                        TrayNo = item.TrayTempNo,
                                        TrayName = item.TrayTempName,
                                        TrayCapacity=item.TrayTempCapacity,
                                        item.Remark,
                                        UseTimes = '0',
                                        Status = "A", //啟用
                                        LaserMarkingStatus='Y',
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy = UserId,
                                        LastModifiedBy = UserId
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

        #region //UpdateTrayUnBind -- 托盤解除綁定 -- Shintokuro 2023.05.28
        public string UpdateTrayUnBind(int TrayId, string TrayNo, string Place)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int TrayIdBase = -1;
                        int TrayBarcodeLogIdBase = -1;
                        string BarcodeNoBase = "";
                        string nulldate = null;
                        string BarcodeStatusBase = "";
                        int BarcodeQtyBase = -1;

                        //TrayId = Convert.ToInt32(TrayNo.Split(',')[1]);
                        //TrayNo = TrayNo.Split(',')[0];

                        for(var i =0;i< TrayNo.Split(',').Count(); i++)
                        {

                            #region //判斷托盤資訊是否正確
                            sql = @"SELECT TOP 1 a.TrayId, a.BarcodeNo,b.BarcodeStatus,b.BarcodeQty
                                FROM MES.Tray a
                                INNER JOIN MES.Barcode b on a.BarcodeNo = b.BarcodeNo
                                WHERE a.CompanyId = @CompanyId
                                AND a.TrayNo = @TrayNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("TrayNo", TrayNo.Split(',')[i]);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【托盤解除綁定】 - 托盤資料錯誤!");
                            foreach (var item in result)
                            {
                                TrayIdBase = item.TrayId;
                                BarcodeNoBase = item.BarcodeNo;
                                BarcodeStatusBase = item.BarcodeStatus;
                                BarcodeQtyBase = item.BarcodeQty;
                                if (Place == "PC")
                                {
                                    if (BarcodeStatusBase != "10") throw new SystemException("【托盤解除綁定】 - 條碼狀態要為出貨狀態才能解除Tray盤綁定");
                                }
                                else if (Place == "Pad")
                                {
                                    if (BarcodeQtyBase != 0) throw new SystemException("【托盤解除綁定】 - 目前條碼數量不為0不會自動解除綁定");
                                }
                            }
                            #endregion

                            #region //撈取Tray盤綁定紀錄ID
                            sql = @"SELECT TOP 1 TrayBarcodeLogId
                                FROM MES.TrayBarcodeLog
                                WHERE TrayId = @TrayId
                                AND BarcodeNo = @BarcodeNoBase
                                AND ReMoveBindDate is null";
                            dynamicParameters.Add("TrayId", TrayIdBase);
                            dynamicParameters.Add("BarcodeNoBase", BarcodeNoBase);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【托盤解除綁定】 - 找不到綁定過往紀錄,請重新確認");
                            foreach (var item in result)
                            {
                                TrayBarcodeLogIdBase = item.TrayBarcodeLogId;
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.Tray SET
                                BarcodeNo = @BarcodeNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayId = @TrayIdBase";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BarcodeNo = nulldate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    TrayIdBase
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.TrayBarcodeLog SET
                                ReMoveBindDate = @LastModifiedDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayBarcodeLogId = @TrayBarcodeLogIdBase";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ReMoveBindDate = LastModifiedDate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    TrayBarcodeLogIdBase
                                });
                            var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                            int rowsAffected1 = sqlConnection.Execute(sql, dynamicParameters);

                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + "1" + " rows affected)"
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

        #region //UpdateKeyenceFlag -- 更新Keyence支援狀態 -- Ann 2023-05-25
        public string UpdateKeyenceFlag(int ParameterId, int MachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷製程機台資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.KeyenceFlag
                                FROM MES.ProcessMachine a 
                                WHERE a.ParameterId = @ParameterId
                                AND a.MachineId = @MachineId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程機台資料錯誤!");
                        #endregion

                        #region //調整為相反狀態
                        string KeyenceFlag = "";
                        foreach (var item in result)
                        {
                            KeyenceFlag = item.KeyenceFlag;
                        }

                        switch (KeyenceFlag)
                        {
                            case "Y":
                                KeyenceFlag = "N";
                                break;
                            case "N":
                                KeyenceFlag = "Y";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ProcessMachine SET
                                KeyenceFlag = @KeyenceFlag,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        var parametersObject = new
                        {
                            KeyenceFlag,
                            LastModifiedDate,
                            LastModifiedBy,
                            ParameterId,
                            MachineId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //UpdateUserEventSetting -- 人員事件設定更新 -- Xuan 2023.07.31
        public string UpdateUserEventSetting(int UserEventItemId, string UserEventItemName)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//人員事件ID是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEventItem
                                WHERE UserEventItemId=@UserEventItemId";
                        dynamicParameters.Add("UserEventItemId", UserEventItemId);                    
                        var UserEventItemIdResult = sqlConnection.Query(sql, dynamicParameters);
                        if (UserEventItemIdResult.Count() != 1) throw new SystemException("【人員事件ID】重複!");
                        #endregion

                        #region//人員事件名稱是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEventItem
                                WHERE UserEventItemName=@UserEventItemName
                                AND CompanyId=@CompanyId";
                        dynamicParameters.Add("UserEventItemName", UserEventItemName);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var UserEventItemNameResult = sqlConnection.Query(sql, dynamicParameters);
                        if (UserEventItemNameResult.Count() != 0) throw new SystemException("【人員事件】重複!");
                        #endregion

                        #region//判斷是否有與另一張table繫結
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEvent
                                WHERE UserEventItemId = @UserEventItemId";
                        dynamicParameters.Add("UserEventItemId", UserEventItemId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0)
                        {

                            #region //更新SQL
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE [MES].[UserEventItem] SET
                                    UserEventItemName=@UserEventItemName,
                                    LastModifiedDate=@LastModifiedDate,
                                    LastModifiedBy=@LastModifiedBy
                                    WHERE UserEventItemId=@UserEventItemId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserEventItemName,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserEventItemId
                            });
                        }
                        else
                        {
                            throw new SystemException("此人員事件已過站使用過!");
                        }

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateMachineEventSetting -- 機台事件設定更新 -- Xuan 2023.08.11
        public string UpdateMachineEventSetting(int MachineId, int MachineEventItemId, string MachineEventName)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                       #region//機台事件ID是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEventItem
                                WHERE MachineEventItemId=@MachineEventItemId";
                        dynamicParameters.Add("MachineEventItemId", MachineEventItemId);
                        var MachineEventItemIdResult = sqlConnection.Query(sql, dynamicParameters);
                        if (MachineEventItemIdResult.Count() != 1) throw new SystemException("【機台事件ID】重複!");
                        #endregion
                        #region //判斷是否有與另一張table繫結
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEvent
                                WHERE MachineEventItemId = @MachineEventItemId";
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0)
                        {

                            if (MachineId <= 0) throw new SystemException("【機台】不能為空!");
                            if (MachineEventName.Length <= 0) throw new SystemException("【事件名稱】不能為空!");
                            if (MachineEventName.Length > 20) throw new SystemException("【事件名稱】長度錯誤!");

                            #region //更新SQL
                            dynamicParameters = new DynamicParameters();

                            sql = @"update [MES].[MachineEventItem]
                                set MachineEventName=@MachineEventName, MachineId=@MachineId, LastModifiedDate=@LastModifiedDate, LastModifiedBy=@LastModifiedBy
                                where MachineEventItemId=@MachineEventItemId";
                            dynamicParameters.AddDynamicParams(
                          new
                          {
                              MachineId,
                              MachineEventName,
                              LastModifiedDate,
                              LastModifiedBy,
                              MachineEventItemId
                          });
                        }
                        else
                        {
                            throw new SystemException("此機台事件已過站使用過!");
                        }

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateProcessEventSetting -- 加工事件設定更新 -- Xuan 2023.08.15
        public string UpdateProcessEventSetting(int ProcessEventItemId, string ProcessEventName)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        
                        if (ProcessEventName.Length <= 0) throw new SystemException("【事件名稱】不能為空!");
                        if (ProcessEventName.Length > 20) throw new SystemException("【事件名稱】長度錯誤!");
                       
                        #region//加工事件ID是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessEventItem
                                WHERE ProcessEventItemId=@ProcessEventItemId";
                        dynamicParameters.Add("ProcessEventItemId", ProcessEventItemId);
                        var ProcessEventItemIdResult = sqlConnection.Query(sql, dynamicParameters);
                        if (ProcessEventItemIdResult.Count() != 1) throw new SystemException("【加工事件ID】重複!");
                        #endregion

                        int rowsAffected = 0;

                        #region //判斷是否有與另一張table繫結
                        sql = @"SELECT TOP 1 1
                                FROM MES.BarcodeProcessEvent
                                WHERE ProcessEventItemId = @ProcessEventItemId";
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0)
                        {
                            #region //更新SQL
                            dynamicParameters = new DynamicParameters();

                            sql = @"update [MES].[ProcessEventItem]
                                set ProcessEventName=@ProcessEventName, LastModifiedDate=@LastModifiedDate, LastModifiedBy=@LastModifiedBy
                                where ProcessEventItemId=@ProcessEventItemId";
                            dynamicParameters.AddDynamicParams(
                          new
                          {
                              ProcessEventName,
                              LastModifiedDate,
                              LastModifiedBy,
                              ProcessEventItemId
                          });
                        }
                        else
                        {
                            throw new SystemException("此加工事件已過站使用過!");
                        }
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateProcessParameterQcitemExcel -- 上傳製程量測參數Excel資料 -- GPAI 2024-01-13
        public string UpdateProcessParameterQcitemExcel(int ParameterId, string ExcelJson)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        List<int> qcitemidList = new List<int>();

                        int rowsAffected = 0;
                        #region//檢核製程是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM MES.ProcessParameter a
                                WHERE a.ParameterId=@ParameterId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("查無製程參數資料,請重新確認");
                        foreach (var item in result)
                        {

                        }
                        #endregion

                        #region//撈取Id
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT QcItemId
                                FROM MES.ProcessParameterQcItem a
                                WHERE a.ParameterId=@ParameterId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                qcitemidList.Add(item.QcItemId);
                            }
                        }
                        #endregion



                        var spreadsheetJson = JObject.Parse(ExcelJson);

                        #region //解析Spreadsheet Data
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {
                            string QcItemNo = item["QcItemNo"] != null ? item["QcItemNo"].ToString() : throw new SystemException("【資料維護不完整】QcItemNo欄位資料不可以為空,請重新確認~~");
                            string QcItemName = item["QcItemName"] != null ? item["QcItemName"].ToString() : throw new SystemException("【資料維護不完整】QcItemName欄位資料不可以為空,請重新確認~~");
                            string QcItemDesc = item["QcItemDesc"] != null ? item["QcItemDesc"].ToString() : "";
                            decimal? DesignValue = item["DesignValue"] != null ? Convert.ToDecimal(item["DesignValue"]) : -9999;
                            decimal? UpperTolerance = item["UpperTolerance"] != null ? Convert.ToDecimal(item["UpperTolerance"]) : -9999;
                            decimal? LowerTolerance = item["LowerTolerance"] != null ? Convert.ToDecimal(item["LowerTolerance"]) : -9999;
                            string Remark = item["Remark"] != null ? item["Remark"].ToString() : "";
                            string BallMark = item["BallMark"] != null ? item["BallMark"].ToString() : "";
                            string Unit = item["Unit"] != null ? item["Unit"].ToString() : "";
                            string MachineDesc = item["MachineDesc"] != null ? item["MachineDesc"].ToString() : "";


                            //var nQcItemNos = jsonData["spreadsheetInfo"].Select(s => s["QcItemNo"]?.ToString().Length == 10 ? s["QcItemNo"]?.ToString().Substring(0, 3) + s["QcItemNo"]?.ToString().Substring(6) : s["QcItemNo"]?.ToString()).ToList();
                            var nQcItemNosd = item["QcItemNo"].ToString().Length == 10 ? item["QcItemNo"]?.ToString().Substring(0, 3) + item["QcItemNo"]?.ToString().Substring(6) : item["QcItemNo"]?.ToString();

                            #region //取得量測儀器
                            decimal? QmmDetailId = null;
                            if (item["QcItemNo"]?.ToString().Length == 10)
                            {
                                string MachineNumber = item["QcItemNo"]?.ToString().Substring(3, 3);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QmmDetailId, a.MachineNumber, b.MachineName, b.MachineNo
                                        FROM  QMS.QmmDetail a
                                        INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                                        INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                                        where a.MachineNumber = @MachineNumber AND c.CompanyId = @CompanyId";
                                dynamicParameters.Add("MachineNumber", MachineNumber);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var QmmDetailIdValue = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var qmmItem in QmmDetailIdValue)
                                {
                                    QmmDetailId = qmmItem.QmmDetailId;
                                }
                            }
                            #endregion

                            if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                            #region //確認項目資料是否正確
                            int QcItemId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM QMS.QcItem a 
                                WHERE a.QcItemNo = @QcItemNo";
                            dynamicParameters.Add("QcItemNo", nQcItemNosd);

                            var QcItemResult = sqlConnection.Query(sql, dynamicParameters);
                            if (!QcItemResult.Any()) throw new SystemException("找不到【" + QcItemName + "(" + nQcItemNosd + ")" + "】量測項目,請重新確認~~");
                            foreach (var item1 in QcItemResult)
                            {
                                QcItemId = item1.QcItemId;
                            }
                            #endregion

                            #region //確認是否已經存在於製程項目表表
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                    FROM MES.ProcessParameterQcItem a 
                                    WHERE a.ParameterId = @ParameterId AND a.QcItemId = @QcItemId
                                    ";
                            dynamicParameters.Add("ParameterId", ParameterId);
                            dynamicParameters.Add("QcItemId", QcItemId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item2 in result)
                                {
                                    if (item2.QcItemId == QcItemId)
                                    {
                                        qcitemidList.Remove(item2.QcItemId);

                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.ProcessParameterQcItem SET
                                                QcItemId = @QcItemId,
                                                DesignValue = @DesignValue,
                                                UpperTolerance = @UpperTolerance,
                                                LowerTolerance = @LowerTolerance,
                                                Remark = @Remark,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy,
                                                QcItemDesc = @QcItemDesc,
                                                BallMark = @BallMark,
                                                Unit = @Unit,
                                                QmmDetailId = @QmmDetailId
                                                WHERE ParameterQcItemId = @ParameterQcItemId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                QcItemId,
                                                DesignValue = DesignValue != -9999 ? DesignValue : null,
                                                UpperTolerance = UpperTolerance != -9999 ? UpperTolerance : null,
                                                LowerTolerance = LowerTolerance != -9999 ? LowerTolerance : null,
                                                Remark,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                QcItemDesc,
                                                BallMark,
                                                Unit,
                                                QmmDetailId,
                                                ParameterQcItemId = item2.ParameterQcItemId
                                            });

                                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                    else
                                    {
                                        throw new SystemException("該【" + QcItemName + "(" + QcItemNo + ")" + "】已經建立,請重新確認~~");
                                    }
                                }
                            }
                            else
                            {
                                #region //INSERT PDM.DfmQiProcess0
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.ProcessParameterQcItem (ParameterId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, QcItemDesc, BallMark, Unit, QmmDetailId)
                                    OUTPUT INSERTED.ParameterQcItemId
                                    VALUES (@ParameterId, @QcItemId, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @QcItemDesc, @BallMark, @Unit, @QmmDetailId)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ParameterId,
                                        QcItemId,
                                        DesignValue = DesignValue != -9999 ? DesignValue : null,
                                        UpperTolerance = UpperTolerance != -9999 ? UpperTolerance : null,
                                        LowerTolerance = LowerTolerance != -9999 ? LowerTolerance : null,
                                        Remark,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy,
                                        QcItemDesc,
                                        BallMark,
                                        Unit,
                                        QmmDetailId
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion
                            }
                            #endregion
                        }
                        #endregion


                        #region //以上傳資料為最新,故要移除沒有再上傳資料中的Bom元件
                        if (qcitemidList.Count() > 0)
                        {
                            foreach (var item in qcitemidList)
                            {
                                #region //刪子table
                                sql = @" DELETE a FROM MES.ProcessParameterQcItem a
                                 INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                 WHERE b.QcItemId = @QcItemId AND a.ParameterId = @ParameterId";
                                dynamicParameters.Add("QcItemId", item);
                                dynamicParameters.Add("ParameterId", ParameterId);


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

        #region //UpdateRoutingItemQcItemExcel -- 上傳途程量測Excel資料 -- GPAI 2024-01-13
        public string UpdateRoutingItemQcItemExcel(int RoutingItemId, string ExcelJson, int ItemProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        List<int> qcitemidList = new List<int>();

                        int rowsAffected = 0;
                        #region//檢核途程是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM MES.RoutingItem a
                                WHERE a.RoutingItemId=@RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("查無製程參數資料,請重新確認");
                        foreach (var item in result)
                        {

                        }
                        #endregion

                        #region//撈取Id
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT QcItemId
                                FROM MES.RoutingItemQcItem a
                                WHERE a.RoutingItemId=@RoutingItemId AND a.ItemProcessId = @ItemProcessId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);
                        dynamicParameters.Add("ItemProcessId", ItemProcessId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                qcitemidList.Add(item.QcItemId);
                            }
                        }
                        #endregion



                        var spreadsheetJson = JObject.Parse(ExcelJson);

                        #region //解析Spreadsheet Data
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {

                            //int? ItemProcessId = item["ItemProcessId"] != null ? Convert.ToInt32(item["ItemProcessId"]) : -1;
                            string QcItemNo = item["QcItemNo"] != null ? item["QcItemNo"].ToString() : throw new SystemException("【資料維護不完整】項目序號序號欄位資料不可以為空,請重新確認~~");
                            //string ProcessAlias = item["ProcessAlias"] != null ? item["ProcessAlias"].ToString() : throw new SystemException("【資料維護不完整】製程欄位資料不可以為空,請重新確認~~");
                            string QcItemName = item["QcItemName"] != null ? item["QcItemName"].ToString() : throw new SystemException("【資料維護不完整】檢測項目欄位資料不可以為空,請重新確認~~");
                            string QcItemDesc = item["QcItemDesc"] != null ? item["QcItemDesc"].ToString() : "";
                            decimal? DesignValue = item["DesignValue"] != null ? Convert.ToDecimal(item["DesignValue"]) : -9999;
                            decimal? UpperTolerance = item["UpperTolerance"] != null ? Convert.ToDecimal(item["UpperTolerance"]) : -9999;
                            decimal? LowerTolerance = item["LowerTolerance"] != null ? Convert.ToDecimal(item["LowerTolerance"]) : -9999;
                            string Remark = item["Remark"] != null ? item["Remark"].ToString() : "";

                            string BallMark = item["BallMark"] != null ? item["BallMark"].ToString() : "";
                            string Unit = item["Unit"] != null ? item["Unit"].ToString() : "";
                            string MachineDesc = item["MachineDesc"] != null ? item["MachineDesc"].ToString() : "";

                            //var nQcItemNos = jsonData["spreadsheetInfo"].Select(s => s["QcItemNo"]?.ToString().Length == 10 ? s["QcItemNo"]?.ToString().Substring(0, 3) + s["QcItemNo"]?.ToString().Substring(6) : s["QcItemNo"]?.ToString()).ToList();
                            var nQcItemNosd = item["QcItemNo"].ToString().Length == 10 ? item["QcItemNo"]?.ToString().Substring(0, 3) + item["QcItemNo"]?.ToString().Substring(6) : item["QcItemNo"]?.ToString();

                            #region //取得量測儀器
                            decimal? QmmDetailId = null;
                            if (item["QcItemNo"]?.ToString().Length == 10)
                            {
                                string MachineNumber = item["QcItemNo"]?.ToString().Substring(3, 3);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QmmDetailId, a.MachineNumber, b.MachineName, b.MachineNo
                                        FROM  QMS.QmmDetail a
                                        INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                                        INNER JOIN MES.WorkShop c on b.ShopId = c.ShopId
                                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                                        where a.MachineNumber = @MachineNumber AND c.CompanyId = @CompanyId";
                                dynamicParameters.Add("MachineNumber", MachineNumber);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var QmmDetailIdValue = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var qmmItem in QmmDetailIdValue)
                                {
                                    QmmDetailId = qmmItem.QmmDetailId;
                                }
                            }
                            #endregion



                            if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                            #region //確認項目資料是否正確
                            int QcItemId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM QMS.QcItem a 
                                WHERE a.QcItemNo = @QcItemNo";
                            dynamicParameters.Add("QcItemNo", nQcItemNosd);

                            var QcItemResult = sqlConnection.Query(sql, dynamicParameters);
                            if (!QcItemResult.Any()) throw new SystemException("找不到【" + QcItemName + "(" + nQcItemNosd + ")" + "】量測項目,請重新確認~~");
                            foreach (var item1 in QcItemResult)
                            {
                                QcItemId = item1.QcItemId;
                            }
                            #endregion

                            #region //確認製程資料是否正確

                            if (ItemProcessId != -1) {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT *
                                FROM MES.RoutingItemProcess  a 
                                WHERE a.ItemProcessId = @ItemProcessId";
                                dynamicParameters.Add("ItemProcessId", ItemProcessId);

                                var RoutingItemProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                if (!RoutingItemProcessResult.Any()) throw new SystemException("找不到該製程,請重新確認~~");
                            }
                           
                            
                            #endregion

                            #region //確認是否已經存在於製程項目表表
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                    FROM MES.RoutingItemQcItem a 
                                    WHERE a.RoutingItemId = @RoutingItemId AND a.QcItemId = @QcItemId AND a.ItemProcessId =  @ItemProcessId
                                    ";
                            dynamicParameters.Add("RoutingItemId", RoutingItemId);
                            dynamicParameters.Add("QcItemId", QcItemId);
                            dynamicParameters.Add("ItemProcessId", ItemProcessId);


                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item2 in result)
                                {
                                    if (item2.QcItemId == QcItemId)
                                    {
                                        qcitemidList.Remove(item2.QcItemId);

                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.RoutingItemQcItem SET
                                                ItemProcessId = @ItemProcessId,
                                                QcItemId = @QcItemId,
                                                DesignValue = @DesignValue,
                                                UpperTolerance = @UpperTolerance,
                                                LowerTolerance = @LowerTolerance,
                                                Remark = @Remark,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy,
                                                QcItemDesc = @QcItemDesc,
                                                BallMark = @BallMark,
                                                Unit = @Unit,
                                                QmmDetailId = @QmmDetailId
                                                WHERE RoutingItemQcItemId = @RoutingItemQcItemId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ItemProcessId,
                                                QcItemId,
                                                DesignValue = DesignValue != -9999 ? DesignValue: null,
                                                UpperTolerance = UpperTolerance != -9999 ? UpperTolerance : null,
                                                LowerTolerance = LowerTolerance != -9999 ? LowerTolerance : null,
                                                Remark,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                QcItemDesc,
                                                BallMark,
                                                Unit,
                                                QmmDetailId,
                                                RoutingItemQcItemId = item2.RoutingItemQcItemId
                                            });

                                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                    else
                                    {
                                        throw new SystemException("該【" + QcItemName + "(" + QcItemNo + ")" + "】已經建立,請重新確認~~");
                                    }
                                }
                            }
                            else
                            {
                                #region //INSERT
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.RoutingItemQcItem (RoutingItemId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, ItemProcessId, QcItemDesc, BallMark, Unit, QmmDetailId)
                                    OUTPUT INSERTED.RoutingItemQcItemId
                                    VALUES (@RoutingItemId, @QcItemId, @DesignValue, @UpperTolerance, @LowerTolerance, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ItemProcessId, @QcItemDesc, @BallMark, @Unit, @QmmDetailId)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoutingItemId,
                                        QcItemId,
                                        DesignValue = DesignValue != -9999 ? DesignValue : null,
                                        UpperTolerance = UpperTolerance != -9999 ? UpperTolerance : null,
                                        LowerTolerance = LowerTolerance != -9999 ? LowerTolerance : null,
                                        Remark,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy,
                                        ItemProcessId,
                                        QcItemDesc,
                                        BallMark,
                                        Unit,
                                        QmmDetailId
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion
                            }
                            #endregion
                        }
                        #endregion


                        #region //以上傳資料為最新,故要移除沒有再上傳資料中的
                        if (qcitemidList.Count() > 0)
                        {
                            foreach (var item in qcitemidList)
                            {
                                #region //刪子table
                                sql = @" DELETE a FROM MES.RoutingItemQcItem a
                                 INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                 WHERE b.QcItemId = @QcItemId AND a.RoutingItemId = @RoutingItemId AND a.ItemProcessId = @ItemProcessId";
                                dynamicParameters.Add("QcItemId", item);
                                dynamicParameters.Add("RoutingItemId", RoutingItemId);
                                dynamicParameters.Add("ItemProcessId", ItemProcessId);


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

        #region //UpdateLensCarrierRing -- 套環資料更新 -- Jean 2025.07.01
        public string UpdateLensCarrierRing(int RingId, string ModelName, string RingName, string Remarks, string HoleCount
            , string RingSpec, string RingCode, string RingShape, string Customer, decimal DailyDemand, decimal SafetyStock)
        {
            try
            {
                if (ModelName.Length <= 0) throw new SystemException("【機種名】不能為空!");
                if (RingCode.Length <= 0) throw new SystemException("【套環編碼】不能為空!");
                if (!int.TryParse(HoleCount, out int holeCountValue)) throw new SystemException("【孔數】必須為整數!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷套環資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.LensCarrierRing
                                WHERE RingId = @RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("套環資料錯誤!");
                        #endregion

                        #region //判斷機種名是否重複
                        sql = @"SELECT TOP 1 1
                                FROM MES.LensCarrierRing
                                WHERE ModelName = @ModelName
                                  AND RingId != @RingId";
                        dynamicParameters.Add("ModelName", ModelName);
                        dynamicParameters.Add("RingId", RingId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【機種名】重複，請重新輸入!");
                        #endregion

                        //#region //判斷套環編號是否重複
                        //sql = @"SELECT TOP 1 1
                        //        FROM MES.LensCarrierRing
                        //        WHERE  RingCode = @RingCode
                        //            AND RingId != @RingId";
                        //dynamicParameters.Add("RingCode", RingCode);
                        //dynamicParameters.Add("RingId", RingId);
                        //var result3 = sqlConnection.Query(sql, dynamicParameters);
                        //if (result3.Count() > 0) throw new SystemException("【套環編號】重複，請重新輸入!");
                        //#endregion

                        #region //判斷機種名+套環編碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM MES.LensCarrierRing
                                WHERE  ModelName = @ModelName AND RingCode = @RingCode
                                   AND RingId != @RingId";
                        dynamicParameters = new DynamicParameters(); // 重置参数避免污染
                        dynamicParameters.Add("ModelName", ModelName);
                        dynamicParameters.Add("RingCode", RingCode);
                        dynamicParameters.Add("RingId", RingId);
                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 0)
                            throw new SystemException("【機種名+套環編碼】組合已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.LensCarrierRing SET
                                ModelName = @ModelName,
                                RingName = @RingName,
                                Remarks = @Remarks,
                                HoleCount = @HoleCount,
                                RingSpec = @RingSpec,
                                RingCode = @RingCode,
                                RingShape = @RingShape,
                                Customer = @Customer,
                                DailyDemand = @DailyDemand,
                                SafetyStock = @SafetyStock,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RingId = @RingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModelName,
                                RingName,
                                Remarks,
                                HoleCount,
                                RingSpec,
                                RingCode,
                                RingShape,
                                Customer,
                                DailyDemand,
                                SafetyStock,
                                LastModifiedDate,
                                LastModifiedBy,
                                RingId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateLensCarrierRingStatus -- 套環資料狀態更新 -- Jean 2025.07.01
        public string UpdateLensCarrierRingStatus(int RingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷套環資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.LensCarrierRing
                                WHERE  RingId = @RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("套環資料錯誤!");

                        #endregion

                        #region //判斷套環資料是否存在库存异动记录
                        sql = @"SELECT TOP 1 1
                                FROM  MES.LensCarrierRing a
                                INNER JOIN MES.RingTransaction b ON a.RingId=b.RingId AND b.Status='A'
                                WHERE  a.RingId = @RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("存在庫存異動記錄，不允許作廢！");

                        #endregion



                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.LensCarrierRing SET
                                Status = 'S',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RingId = @RingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                RingId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateRingTransaction -- 套環庫存異動更新 -- Jean 2025.07.07
        public string UpdateRingTransaction(int RingTransId, int RingId, string TransType, int Quantity)
        {
            try
            {
                if (TransType.Length <= 0) throw new SystemException("【異動類型】不能為空!");
                if (Quantity <= 0) throw new SystemException("【數量】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷套環庫存異動資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.RingTransaction
                                WHERE RingTransId = @RingTransId";
                        dynamicParameters.Add("RingTransId", RingTransId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("套環存異動資料錯誤!");
                        #endregion

                        #region //判斷新增时套环库存是否出现负库存情况
                        sql = @"SELECT x.Quantity
                                FROM (
                                    SELECT a.RingId,ISNULL(SUM(c.Quantity),0)Quantity
                                    FROM  MES.LensCarrierRing a
                                    LEFT JOIN (SELECT e.RingId,(CASE WHEN e.TransType='IN' THEN e.Quantity WHEN e.TransType='OUT' THEN -e.Quantity ELSE 0 END) Quantity
			                                    FROM MES.RingTransaction e
			                                    WHERE e.Status='A' AND e.RingTransId != @RingTransId
		                                    )c ON c.RingId = a.RingId
                                    INNER JOIN BAS.[User] d on a.CreateBy = d.UserId
                                    WHERE 1=1
                                    GROUP BY a.RingId
                                    )x
                                WHERE x.RingId=@RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var currentStock = sqlConnection.QuerySingleOrDefault<int?>(sql, dynamicParameters) ?? 0;

                        int newStock = TransType == "IN"
                            ? currentStock + Quantity
                            : currentStock - Quantity;

                        // 检查是否会导致负库存
                        if (newStock < 0)
                        {
                            throw new SystemException($"當前庫存(不含此記錄): {currentStock}，庫存不足，此操作會導致負庫存! ");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RingTransaction SET
                                TransType = @TransType,
                                Quantity = @Quantity,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RingTransId = @RingTransId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransType,
                                Quantity,
                                LastModifiedDate,
                                LastModifiedBy,
                                RingTransId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateRingTransactionStatus -- 套環庫存異動狀態更新 -- Jean 2025.07.07
        public string UpdateRingTransactionStatus(int RingTransId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷套環庫存異動資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.RingTransaction
                                WHERE  RingTransId = @RingTransId";
                        dynamicParameters.Add("RingTransId", RingTransId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("套環庫存異動資料錯誤!");

                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.RingTransaction SET
                                Status = 'S',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RingTransId = @RingTransId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                RingTransId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #endregion

        #region //Delete
        #region //DeleteProdModeShift -- 生產模式班次資料刪除 -- Shintokru 2022.06.22
        public string DeleteProdModeShift(int ModeShiftId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷生產模式班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdModeShift
                                WHERE ModeShiftId = @ModeShiftId";
                        dynamicParameters.Add("ModeShiftId", ModeShiftId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產模式班次資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除SQL - 主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ProdModeShift
                                WHERE ModeShiftId = @ModeShiftId";
                        dynamicParameters.Add("ModeShiftId", ModeShiftId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteMachineAsset -- 機台資產編號刪除 -- Ted 2022.06.10
        public string DeleteMachineAsset(int MachineId, string AssetNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine a
                                INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId
                                WHERE a.MachineId = @MachineId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        #region //機台資產編號是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineAsset
                                WHERE MachineId = @MachineId
                                AND AssetNo = @AssetNo";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("AssetNo", AssetNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("機台資產編號錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.MachineAsset
                                WHERE MachineId = @MachineId
                                AND AssetNo = @AssetNo";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("AssetNo", AssetNo);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDeviceMachine -- 裝置機台綁定刪除 -- Ted 2022.06.15
        public string DeleteDeviceMachine(int DeviceId, int MachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.DeviceMachine
                                WHERE DeviceId = @DeviceId";
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        #region //機台資產編號是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.DeviceMachine
                                WHERE DeviceId = @DeviceId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("DeviceId", DeviceId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("機台資產編號錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.DeviceMachine
                                WHERE DeviceId = @DeviceId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("DeviceId", DeviceId);
                        dynamicParameters.Add("MachineId", MachineId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDevice -- 裝置刪除 -- Ted 2022.06.21
        public string DeleteDevice(int DeviceId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceId = @DeviceId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceId", DeviceId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.Device
                                WHERE CompanyId = @CompanyId
                                AND DeviceId = @DeviceId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DeviceId", DeviceId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteProcess -- 製程刪除 -- Shintokru 2022.06.21
        public string DeleteProcess(int ProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE CompanyId = @CompanyId
                                AND ProcessId = @ProcessId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ProcessId", ProcessId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                    
                        #region //刪除SQL - 主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.Process
                                WHERE ProcessId = @ProcessId
                                AND CompanyId = CompanyId";
                        dynamicParameters.Add("ProcessId", ProcessId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteProdUnit -- 生產單元資料刪除 -- Shintokru 2022.06.23
        public string DeleteProdUnit(int UnitId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷生產單元資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdUnit
                                WHERE UnitId = @UnitId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("UnitId", UnitId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("生產單元資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除SQL - 主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ProdUnit
                                WHERE UnitId = @UnitId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("UnitId", UnitId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteWarehouseLocation -- 裝置機台綁定刪除 -- Ted 2022.07.05
        public string DeleteWarehouseLocation(int WarehouseId, int LocationId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.WarehouseLocation
                                WHERE WarehouseId = @WarehouseId";
                        dynamicParameters.Add("WarehouseId", WarehouseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        #region //機台資產編號是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.WarehouseLocation
                                WHERE WarehouseId = @WarehouseId
                                AND LocationId = @LocationId";
                        dynamicParameters.Add("WarehouseId", WarehouseId);
                        dynamicParameters.Add("LocationId", LocationId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("機台資產編號錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.WarehouseLocation
                                WHERE WarehouseId = @WarehouseId
                                AND LocationId = @LocationId";
                        dynamicParameters.Add("WarehouseId", WarehouseId);
                        dynamicParameters.Add("LocationId", LocationId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteProcessMachine -- 製程機台資料刪除 -- Ted 2022.07.06
        public string DeleteProcessMachine(int ParameterId, int MachineId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷該製程下機台是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessMachine
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("MachineId", MachineId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ProcessMachine
                                WHERE ParameterId = @ParameterId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("MachineId", MachineId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteProcessProductionUnit -- 製程生產單元刪除 -- Ted 2022.07.06
        public string DeleteProcessProductionUnit(int ParameterId, int UnitId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷該製程下機台是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessProductionUnit
                                WHERE ParameterId = @ParameterId
                                AND UnitId = @UnitId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("UnitId", UnitId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.ProcessProductionUnit
                                WHERE ParameterId = @ParameterId
                                AND UnitId = @UnitId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        dynamicParameters.Add("UnitId", UnitId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteTray -- 托盤刪除 -- Shintokuro 2023.02.04
        public string DeleteTray(int TrayId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷托盤資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.Tray
                                WHERE CompanyId = @CompanyId
                                AND TrayId = @TrayId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("TrayId", TrayId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【托盤管理】 - 托盤資料錯誤!");
                        #endregion

                        #region //判斷托盤是否曾經有綁定條碼
                        sql = @"SELECT TOP 1 1
                                FROM MES.TrayBarcodeLog
                                WHERE TrayId = @TrayId";
                        dynamicParameters.Add("TrayId", TrayId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【托盤管理】 - 托盤資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.Tray
                                WHERE CompanyId = @CompanyId
                                AND TrayId = @TrayId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("TrayId", TrayId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteTrayBatch -- 托盤刪除(批量) -- Shintokuro 2023.02.04
        public string DeleteTrayBatch(string TrayListY)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;


                        foreach (var TrayId in TrayListY.Split(','))
                        {
                            #region //判斷托盤資料是否正確
                            sql = @"SELECT TOP 1 1
                                    FROM MES.Tray
                                    WHERE CompanyId = @CompanyId
                                    AND TrayId = @TrayId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("TrayId", TrayId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【托盤管理】 - 托盤不存在,資料錯誤!");
                            #endregion

                            #region //判斷托盤是否曾經有綁定條碼
                            sql = @"SELECT TOP 1 1
                                    FROM MES.TrayBarcodeLog
                                    WHERE TrayId = @TrayId";
                            dynamicParameters.Add("TrayId", TrayId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【托盤管理】 - 托盤曾經綁定過,不能刪除!");
                            #endregion


                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.Tray
                                    WHERE CompanyId = @CompanyId
                                    AND TrayId = @TrayId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("TrayId", TrayId);

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

        #region //DeleteUserEventSetting -- 人員事件資料刪除 -- Xuan 2023.08.07
        public string DeleteUserEventSetting(int UserEventItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷事件資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEventItem
                                WHERE UserEventItemId = @UserEventItemId";
                        dynamicParameters = new DynamicParameters();                       
                        dynamicParameters.Add("UserEventItemId", UserEventItemId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1)
                        {
                            throw new SystemException("人員事件資料刪除錯誤!");
                        }
                        #endregion

                        #region //判斷是否有與另一張table繫結
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEvent
                                WHERE UserEventItemId = @UserEventItemId";
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0 )
                        {
                            int rowsAffected = 0;

                            #region //刪除SQL - 主要table
                            sql = @"DELETE FROM MES.UserEventItem
                                    WHERE UserEventItemId = @UserEventItemId";
                            dynamicParameters.Add("UserEventItemId", UserEventItemId);
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            if(rowsAffected!=1) throw new SystemException("人員事件資料刪除錯誤!");
                            #endregion

                            #region //Response
                            var jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion

                            transactionScope.Complete();
                        }
                        else {
                            throw new SystemException("此人員事件已過站使用過!");
                        }                     
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

        #region //DeleteMachineEventSetting -- 機台事件資料刪除 -- Xuan 2023.08.11
        public string DeleteMachineEventSetting(int MachineEventItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷事件資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEventItem
                                WHERE MachineEventItemId = @MachineEventItemId
                                AND CompanyId = @CompanyId";
                        dynamicParameters = new DynamicParameters();
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MachineEventItemId", MachineEventItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("事件資料錯誤!");

                        #region //判斷是否有與另一張table繫結
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEvent
                                WHERE MachineEventItemId = @MachineEventItemId";
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0)
                        {

                            int rowsAffected = 0;

                            #region //刪除SQL - 主要table
                            sql = @"DELETE FROM MES.MachineEventItem
                                    WHERE MachineEventItemId = @MachineEventItemId";

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            if (rowsAffected != 1) throw new SystemException("機台事件資料刪除錯誤!");
                            #endregion

                            #region //Response
                            var jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion

                            transactionScope.Complete();
                        }
                        else
                        {
                            throw new SystemException("此機台事件已過站使用過!");
                        }
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

        #region //DeleteProcessEventSetting -- 加工事件資料刪除 -- Xuan 2023.08.15
        public string DeleteProcessEventSetting(int ProcessEventItemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷事件資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProcessEventItem
                                WHERE ProcessEventItemId = @ProcessEventItemId";
                        dynamicParameters = new DynamicParameters();
                        dynamicParameters.Add("ProcessEventItemId", ProcessEventItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("事件資料錯誤!");

                        #region //判斷是否有與另一張table繫結
                        sql = @"SELECT TOP 1 1
                                FROM MES.BarcodeProcessEvent
                                WHERE ProcessEventItemId = @ProcessEventItemId";
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0)
                        {

                            int rowsAffected = 0;

                            #region //刪除SQL - 主要table
                            sql = @"DELETE FROM MES.ProcessEventItem
                                    WHERE ProcessEventItemId = @ProcessEventItemId";

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            if (rowsAffected != 1) throw new SystemException("加工事件資料刪除錯誤!");
                            #endregion

                            #region //Response
                            var jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion

                            transactionScope.Complete();
                        }
                        else
                        {
                            throw new SystemException("此機台事件已過站使用過!");
                        }
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

        #region //DeleteAllProcessParameterQcitem -- 刪除全部製程量測項目 -- GPAI 2024-01-15
        public string DeleteAllProcessParameterQcitem(int ParameterId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核製程是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM MES.ProcessParameter a
                                WHERE a.ParameterId=@ParameterId";
                        dynamicParameters.Add("ParameterId", ParameterId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【BOM單頭】查無資料,請重新確認");
                        foreach (var item in result)
                        {
                        }
                        #endregion


                        #region //刪除PDM.BomDetail
                        sql = @"DELETE MES.ProcessParameterQcItem
                                WHERE ParameterId = @ParameterId";
                        dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ParameterId
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

        #region //DeleteAllRoutingItemQcItem -- 刪除全部途程量測項目 -- GPAI 2024-01-15
        public string DeleteAllRoutingItemQcItem(int RoutingItemId, int ItemProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核製程是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM MES.RoutingItem a
                                WHERE a.RoutingItemId=@RoutingItemId";
                        dynamicParameters.Add("RoutingItemId", RoutingItemId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【BOM單頭】查無資料,請重新確認");
                        foreach (var item in result)
                        {
                        }
                        #endregion


                        #region //刪除
                        sql = @"DELETE MES.RoutingItemQcItem
                                WHERE RoutingItemId = @RoutingItemId AND ItemProcessId = @ItemProcessId";
                        dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RoutingItemId,
                                    ItemProcessId
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

        #region //DeleteLensCarrierRing -- 套環資料刪除 -- Jean 2025.07.01
        public string DeleteLensCarrierRing(int RingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.LensCarrierRing
                                WHERE  RingId = @RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("套環資料錯誤!");
                        #endregion

                        #region //判斷是否存在套環庫存異動資料
                        sql = @"SELECT TOP 1 1
                            FROM MES.LensCarrierRing a
                            INNER JOIN MES.RingTransaction b ON a.RingId=b.RingId AND b.Status='A'
                            WHERE a.RingId=@RingId";
                        dynamicParameters.Add("RingId", RingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("已存在套環庫存異動資料，不可刪除!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.LensCarrierRing
                                WHERE  RingId = @RingId";
                        dynamicParameters.Add("RingId", RingId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRingTransaction -- 套環庫存異動資料刪除 -- Jean 2025.07.07
        public string DeleteRingTransaction(int RingTransId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷庫存異動資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.RingTransaction
                                WHERE  RingTransId = @RingTransId";
                        dynamicParameters.Add("RingTransId", RingTransId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("套環庫存異動資料錯誤!");
                        #endregion
                        
                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.RingTransaction
                                WHERE  RingTransId = @RingTransId";
                        dynamicParameters.Add("RingTransId", RingTransId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //FOR EIP API
        #region //GetProdModeEIP -- 取得生產模式 -- GPAI 2024.03.15
        public string GetProdModeEIP(int ModeId, string ModeNo, string ModeName, string Status, string BarcodeCtrl, string ScrapRegister
            , string OrderBy, int PageIndex, int PageSize, int[] CustomerIds)
        {
            try
            {
                if (CustomerIds == null) throw new SystemException("客戶資料尚未綁定");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得客戶公司資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.CompanyId
                            FROM SCM.Customer a
                            WHERE a.CustomerId IN @CustomerIds";
                    dynamicParameters.Add("CustomerIds", CustomerIds);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    #endregion

                    #region //取得生產模式資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ModeId";
                    //sqlQuery.distinct = true;
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.CompanyId, a.ModeNo, a.ModeName, a.ModeDesc, a.Status ,a.BarcodeCtrl, a.ScrapRegister
                           ,a.Source ,a.PVTQCFlag ,a.NgToBarcode ,a.TrayBarcode ,a.LotStatus ,a.OutputBarcodeFlag ,a.MrType ,a.OQcCheckType
                           ,(a.ModeNo + '-' + a.ModeName) txtProdModeWithText";
                    sqlQuery.mainTables =
                        @"FROM MES.ProdMode a";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @" AND a.CompanyId IN @CompanyIds";
                    dynamicParameters.Add("CompanyIds", result.Select(s => s.CompanyId));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ModeId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeNo", @" AND a.ModeNo LIKE '%' + @ModeNo + '%'", ModeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeName", @" AND (a.ModeName LIKE '%' + @ModeName + '%' OR a.ModeDesc LIKE '%' + @ModeName + '%')", ModeName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (BarcodeCtrl.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeCtrl", @" AND a.BarcodeCtrl IN @BarcodeCtrl", BarcodeCtrl.Split(','));
                    if (ScrapRegister.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ScrapRegister", @" AND a.ScrapRegister IN @ScrapRegister", ScrapRegister.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ModeId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
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

        #region //GetProcessEIP -- 取得製程資料 --  GPAI 2024.03.15
        public string GetProcessEIP(int ProcessId, string ProcessNo, string ProcessName, string Status
            , string OrderBy, int PageIndex, int PageSize, int[] CustomerIds)
        {
            try
            {
                if (CustomerIds == null) throw new SystemException("客戶資料尚未綁定");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    //sql = @"select DISTINCT a.ProcessId, a.ProcessNo, a.ProcessName, a.ProcessDesc, a.Status, (a.ProcessNo + '-' + a.ProcessName) ProcessWithText
			              //FROM MES.Process a 
                    //       INNER JOIN BAS.Company b on a.CompanyId = b.CompanyId
                    //       INNER JOIN SCM.Customer c on b.CompanyId = c.CompanyId
                    //       where c.CustomerId IN @CustomerId";
                    //dynamicParameters.Add("CustomerId", CustomerIds);

                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessId", @" AND a.ProcessId = @ProcessId", ProcessId);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessNo", @" AND a.ProcessNo LIKE '%' + @ProcessNo + '%'", ProcessNo);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessName", @" AND a.ProcessName LIKE '%' + @ProcessName + '%'", ProcessName);
                    //if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    //var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    sqlQuery.mainKey = "a.ProcessId, c.CustomerId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProcessNo, a.ProcessName, a.ProcessDesc, a.Status, (a.ProcessNo + '-' + a.ProcessName) ProcessWithText";
                    sqlQuery.mainTables =
                        @"FROM MES.Process a 
                        INNER JOIN BAS.Company b on a.CompanyId = b.CompanyId
                        INNER JOIN SCM.Customer c on b.CompanyId = c.CompanyId";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CustomerId IN @CustomerId";
                    dynamicParameters.Add("CustomerId", CustomerIds);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND a.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessNo", @" AND a.ProcessNo LIKE '%' + @ProcessNo + '%'", ProcessNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessName", @" AND a.ProcessName LIKE '%' + @ProcessName + '%'", ProcessName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProcessId";
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


        #region//AI 晶彩雷刻機

        #region//apiAddTrayAsync  -- 托盤資料新增 (非同步版本) -- Fixed Transaction
        public async Task<string> apiAddTrayAsync(string Company, string UserNo, string TrayPrefix, string TrayName,
        int TrayCapacity, int SerialNumber, int Fabrication, int Serial, string SuffixCode, string Remark)
        {
            try
            {
                // 調試日誌 - 記錄輸入參數
                logger.Info($"apiAddTrayAsync 開始執行 - Company: {Company}, UserNo: {UserNo}, TrayPrefix: {TrayPrefix}, TrayName: {TrayName}, TrayCapacity: {TrayCapacity}, SerialNumber: {SerialNumber}, Fabrication: {Fabrication}, Serial: {Serial}, SuffixCode: {SuffixCode}");

                // 參數驗證
                ValidateAddTrayParameters(TrayPrefix, TrayName, TrayCapacity, SerialNumber, Fabrication, Serial);

                using (var transactionScope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
                {
                    using (var sqlConnection = new SqlConnection(GetOptimizedConnectionString()))
                    {
                        await sqlConnection.OpenAsync();

                        #region //確認公司別DB
                        var companyInfo = await GetCompanyInfoAsync(sqlConnection, Company);
                        if (companyInfo == null) throw new SystemException("公司別錯誤!");
                        logger.Info($"公司資訊取得成功 - CompanyId: {companyInfo.CompanyId}");
                        #endregion

                        #region //取得User資料
                        var userId = await GetUserIdAsync(sqlConnection, UserNo);
                        if (userId <= 0) throw new SystemException("使用者資料錯誤!!");
                        logger.Info($"使用者資訊取得成功 - UserId: {userId}");
                        #endregion

                        #region //產生托盤編號並批量插入
                        var trayNumbers = await GenerateTrayNumbersAsync(sqlConnection, TrayPrefix, Serial, SerialNumber,
                            Fabrication, SuffixCode);

                        logger.Info($"產生托盤編號數量: {trayNumbers.Count}");
                        if (trayNumbers.Count > 0)
                        {
                            logger.Info($"托盤編號列表: {string.Join(", ", trayNumbers)}");
                        }

                        var insertedTrays = await BatchInsertTraysAsync(sqlConnection, trayNumbers, TrayName,
                            TrayCapacity, Remark, companyInfo.CompanyId, userId);

                        logger.Info($"成功插入托盤數量: {insertedTrays.Count}");
                        #endregion

                        #region //Response - 直接回傳API文件要求的格式
                        var resultArray = new JArray();
                        for (int i = 0; i < insertedTrays.Count; i++)
                        {
                            var tray = insertedTrays[i];
                            resultArray.Add(new JObject { ["TrayNo"] = tray.TrayNo.ToString() });
                            logger.Info($"新增托盤: {tray.TrayNo}");
                        }

                        var responseObj = new JObject
                        {
                            ["status"] = "success",
                            ["msg"] = "OK",
                            ["result"] = resultArray
                        };

                        logger.Info($"最終回傳格式: {responseObj}");

                        // 必須在這裡提交交易，在 using 區塊結束前
                        transactionScope.Complete();
                        logger.Info("交易已提交");

                        return responseObj.ToString();
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                #region //Response - 符合API文件格式
                var errorObj = new JObject
                {
                    ["status"] = "error",
                    ["msg"] = e.Message
                };
                #endregion               
                return errorObj.ToString();
            }
        }

        private async Task<List<string>> GenerateTrayNumbersAsync(SqlConnection connection, string trayPrefix,
            int serial, int serialNumber, int fabrication, string suffixCode)
        {
            logger.Info($"GenerateTrayNumbersAsync 開始 - TrayPrefix: {trayPrefix}, Serial: {serial}, SerialNumber: {serialNumber}, Fabrication: {fabrication}, SuffixCode: {suffixCode}");

            var serialNumMax = await GetMaxSerialNumberAsync(connection, trayPrefix, serial);
            logger.Info($"目前最大流水號: {serialNumMax}");

            var productNum = 0;

            // 修正邏輯：根據參數決定產生數量
            if (fabrication > 0)
            {
                // 生產數量模式：從最大號碼+1開始，產生 fabrication 個
                productNum = serialNumMax + fabrication;
                logger.Info($"生產數量模式 - 將產生從 {serialNumMax + 1} 到 {productNum} 的托盤");
            }
            else if (serialNumber > 0)
            {
                // 流水編號模式：產生到指定編號
                if (serialNumMax >= serialNumber)
                    throw new SystemException("【托盤管理】- 目前設定流水編號系統已經存在!");
                productNum = serialNumber;
                logger.Info($"流水編號模式 - 將產生從 {serialNumMax + 1} 到 {productNum} 的托盤");
            }
            else
            {
                // 如果都沒有設定，預設產生1個
                productNum = serialNumMax + 1;
                logger.Info($"預設模式 - 將產生1個托盤，編號為 {productNum}");
            }

            var trayNumbers = new List<string>();
            var serialZero = new string('0', serial);
            var pattern = @"^[A-Z0-9\\-]*$";

            if (!Regex.Match(trayPrefix, pattern).Success)
                throw new SystemException("【托盤管理】- 托盤編號只能由大寫英文和數字和【-】組成");

            for (int i = serialNumMax + 1; i <= productNum; i++)
            {
                var serialNum = i.ToString(serialZero);
                if (serialZero.Length != serialNum.Length)
                {
                    logger.Error($"流水碼長度錯誤 - 期望長度: {serialZero.Length}, 實際長度: {serialNum.Length}, 數字: {i}");
                    throw new SystemException("【托盤管理】- 已達到該組合流水碼最大號了!");
                }

                var trayNo = suffixCode == "Y" ? $"{trayPrefix}-{serialNum}" : $"{trayPrefix}{serialNum}";
                logger.Info($"準備驗證托盤編號: {trayNo}");

                await ValidateTrayNumberAsync(connection, trayNo);
                trayNumbers.Add(trayNo);
                logger.Info($"托盤編號驗證通過並加入列表: {trayNo}");
            }

            logger.Info($"GenerateTrayNumbersAsync 完成 - 總共產生 {trayNumbers.Count} 個托盤編號");
            return trayNumbers;
        }

        private async Task<int> GetMaxSerialNumberAsync(SqlConnection connection, string trayPrefix, int serial)
        {
            var serialPattern = new string('_', serial);
            logger.Info($"查詢最大流水號 - TrayPrefix: {trayPrefix}, SerialPattern: {serialPattern}");

            var sql = @"
                SELECT TOP 1 TrayNo
                FROM (
                    SELECT TrayNo FROM MES.Tray WHERE TrayNo LIKE @TrayPrefix + @SerialPattern
                    UNION ALL
                    SELECT TrayTempNo AS TrayNo FROM MES.TrayTemp WHERE TrayTempNo LIKE @TrayPrefix + @SerialPattern
                ) AS CombinedTray 
                ORDER BY TrayNo DESC";

            var parameters = new DynamicParameters();
            parameters.Add("TrayPrefix", trayPrefix);
            parameters.Add("SerialPattern", serialPattern);

            var result = await connection.QueryFirstOrDefaultAsync<string>(sql, parameters);
            logger.Info($"查詢到的最新托盤編號: {result ?? "無"}");

            if (string.IsNullOrEmpty(result))
            {
                logger.Info("沒有找到現有托盤，從0開始");
                return 0;
            }

            var maxSerial = Convert.ToInt32(result.Substring(result.Length - serial));
            logger.Info($"解析出的最大流水號: {maxSerial}");
            return maxSerial;
        }

        private async Task ValidateTrayNumberAsync(SqlConnection connection, string trayNo)
        {
            var validationQueries = new[]
            {
                "SELECT TOP 1 1 FROM MES.Tray WHERE TrayNo = @TrayNo",
                "SELECT TOP 1 1 FROM MES.Barcode WHERE BarcodeNo = @TrayNo",
                "SELECT TOP 1 1 FROM MES.BarcodePrint WHERE BarcodeNo = @TrayNo"
            };

            var parameters = new DynamicParameters();
            parameters.Add("TrayNo", trayNo);

            for (int i = 0; i < validationQueries.Length; i++)
            {
                var query = validationQueries[i];
                var exists = await connection.QueryFirstOrDefaultAsync<int?>(query, parameters);
                if (exists.HasValue)
                {
                    var tableName = query.Contains("Tray WHERE") ? "托盤" :
                                   query.Contains("Barcode WHERE") ? "Barcode" : "BarcodePrint";
                    logger.Error($"托盤編號重複 - {trayNo} 已存在於 {tableName} 表");
                    throw new SystemException($"【托盤管理】- 托盤編號已存在{tableName}，請重新輸入!");
                }
            }
            logger.Info($"托盤編號驗證通過: {trayNo}");
        }

        // 其他輔助方法保持不變
        private void ValidateAddTrayParameters(string trayPrefix, string trayName, int trayCapacity,
            int serialNumber, int fabrication, int serial)
        {
            if (string.IsNullOrEmpty(trayPrefix)) throw new SystemException("【托盤前綴】不能為空!");
            if (trayPrefix.Length > 100) throw new SystemException("【托盤前綴】長度錯誤!");
            if (string.IsNullOrEmpty(trayName)) throw new SystemException("【托盤名稱】不能為空!");
            if (trayName.Length > 100) throw new SystemException("【托盤名稱】長度錯誤!");
            if (trayCapacity < 0) throw new SystemException("【托盤容量】不能為小於0!");
            if (serial <= 0) throw new SystemException("【流水碼數】不能為小於0!");

            if (serialNumber > 0 && fabrication != 0)
                throw new SystemException("【流水編號模式】不須Key欲生產托盤數量!");
            if (fabrication > 0)
            {
                if (serialNumber != 0) throw new SystemException("【欲生產托盤數量模式】不須Key流水編號!");
                if (fabrication > 5000) throw new SystemException("【欲生產托盤數量】一次生產最多不能超過5000");
            }
        }

        private async Task<dynamic> GetCompanyInfoAsync(SqlConnection connection, string company)
        {
            var sql = @"SELECT ErpNo, ErpDb, CompanyId FROM BAS.Company WHERE CompanyNo = @CompanyNo";
            var parameters = new DynamicParameters();
            parameters.Add("CompanyNo", company);

            var result = await connection.QueryFirstOrDefaultAsync(sql, parameters);
            return result;
        }

        private async Task<int> GetUserIdAsync(SqlConnection connection, string userNo)
        {
            var sql = @"SELECT UserId FROM BAS.[User] WHERE UserNo = @UserNo";
            var parameters = new DynamicParameters();
            parameters.Add("UserNo", userNo);

            var result = await connection.QueryFirstOrDefaultAsync<int?>(sql, parameters);
            return result ?? -1;
        }

        private async Task<List<dynamic>> BatchInsertTraysAsync(SqlConnection connection, List<string> trayNumbers,
            string trayName, int trayCapacity, string remark, int companyId, int userId)
        {
            var insertedTrays = new List<dynamic>();
            var batchSize = 100; // 批量處理大小

            for (int i = 0; i < trayNumbers.Count; i += batchSize)
            {
                var batch = new List<string>();
                var endIndex = Math.Min(i + batchSize, trayNumbers.Count);

                for (int j = i; j < endIndex; j++)
                {
                    batch.Add(trayNumbers[j]);
                }

                var batchResult = await InsertTrayBatchAsync(connection, batch, trayName, trayCapacity,
                    remark, companyId, userId);

                for (int k = 0; k < batchResult.Count; k++)
                {
                    insertedTrays.Add(batchResult[k]);
                }
            }

            return insertedTrays;
        }

        private async Task<List<dynamic>> InsertTrayBatchAsync(SqlConnection connection, List<string> trayNumbers,
            string trayName, int trayCapacity, string remark, int companyId, int userId)
        {
            var sql = @"
                    INSERT INTO MES.TrayTemp (CompanyId, TrayTempNo, TrayTempName, TrayTempCapacity, Remark, Status, 
                                              CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                    OUTPUT INSERTED.TrayTempNo AS TrayNo
                    VALUES (@CompanyId, @TrayNo, @TrayName, @TrayCapacity, @Remark, @Status, 
                            @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

            var results = new List<dynamic>();
            var now = DateTime.Now;

            for (int i = 0; i < trayNumbers.Count; i++)
            {
                var trayNo = trayNumbers[i];
                var parameters = new DynamicParameters();
                parameters.Add("CompanyId", companyId);
                parameters.Add("TrayNo", trayNo);
                parameters.Add("TrayName", trayName);
                parameters.Add("TrayCapacity", trayCapacity);
                parameters.Add("Remark", remark);
                parameters.Add("Status", "A");
                parameters.Add("CreateDate", now);
                parameters.Add("LastModifiedDate", now);
                parameters.Add("CreateBy", userId);
                parameters.Add("LastModifiedBy", userId);

                logger.Info($"準備插入托盤: {trayNo}");
                var result = await connection.QueryFirstOrDefaultAsync(sql, parameters);
                if (result != null)
                {
                    results.Add(result);
                    logger.Info($"成功插入托盤: {result.TrayNo}");
                }
                else
                {
                    logger.Error($"托盤插入失敗: {trayNo}");
                }
            }

            return results;
        }

        private string GetOptimizedConnectionString()
        {
            // 優化連接字串，增加timeout和重試參數
            var builder = new SqlConnectionStringBuilder(MainConnectionStrings)
            {
                ConnectTimeout = 60,
                Pooling = true,
                MaxPoolSize = 100,
                MinPoolSize = 5
            };

            // 檢查是否支援重試設定 (SQL Server 2014+)
            try
            {
                builder.ConnectRetryCount = 3;
                builder.ConnectRetryInterval = 2;
            }
            catch
            {
                // 如果不支援重試設定，忽略這個錯誤
            }

            return builder.ConnectionString;
        }
        #endregion

        #region //UpdateLaserEngravedRingCodeAsync -- 自動套環雷刻機 托盤狀態更新 (非同步版本) -- Optimized 2024.12.04
        public async Task<string> UpdateLaserEngravedRingCodeAsync(string Company, string UserNo, string TrayNo)
        {
            try
            {
                // 調試日誌 - 記錄輸入參數
                logger.Info($"UpdateLaserEngravedRingCodeAsync 開始執行 - Company: {Company}, UserNo: {UserNo}, TrayNo: {TrayNo}");

                using (var transactionScope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
                {
                    using (var sqlConnection = new SqlConnection(GetOptimizedConnectionString()))
                    {
                        await sqlConnection.OpenAsync();

                        #region //確認公司別DB
                        var companyInfo = await GetCompanyInfoAsync(sqlConnection, Company);
                        if (companyInfo == null) throw new SystemException("公司別錯誤!");
                        logger.Info($"公司資訊取得成功 - CompanyId: {companyInfo.CompanyId}");
                        #endregion

                        #region //取得User資料
                        var userId = await GetValidUserIdAsync(sqlConnection, UserNo);
                        if (userId <= 0) throw new SystemException("使用者資料錯誤!!");
                        logger.Info($"使用者資訊取得成功 - UserId: {userId}");
                        #endregion

                        #region//查詢並轉移 TrayTemp 到 Tray
                        var rowsAffected = await TransferTrayTempToTrayAsync(sqlConnection, TrayNo,
                            companyInfo.CompanyId, userId);
                        logger.Info($"轉移托盤資料完成 - 影響行數: {rowsAffected}");
                        #endregion

                        if (rowsAffected==0) throw new SystemException("未正常刷取條碼");
                        #region //Response - 符合API文件格式
                        var responseObj = new JObject
                        {
                            ["status"] = "success",
                            ["msg"] = $"({rowsAffected} rows affected)"
                        };

                        logger.Info($"API02最終回傳格式: {responseObj}");

                        // 必須在這裡提交交易，在 using 區塊結束前
                        transactionScope.Complete();
                        logger.Info("交易已提交");

                        return responseObj.ToString();
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                #region //Response - 符合API文件格式
                var errorObj = new JObject
                {
                    ["status"] = "error",
                    ["msg"] = e.Message
                };
                #endregion
                               
                return errorObj.ToString();
            }
        }

        private async Task<int> GetValidUserIdAsync(SqlConnection connection, string userNo)
        {
            var sql = @"SELECT UserId FROM BAS.[User] WHERE UserNo = @UserNo AND UserId != 216";
            var parameters = new DynamicParameters();
            parameters.Add("UserNo", userNo);

            var result = await connection.QueryFirstOrDefaultAsync<int?>(sql, parameters);
            return result ?? -1;
        }

        private async Task<int> TransferTrayTempToTrayAsync(SqlConnection connection, string trayNo,
            int companyId, int userId)
        {
            var rowsAffected = 0;
            int? nullData = null;
            // 查詢 TrayTemp 資料
            var selectSql = @"
                SELECT CompanyId, TrayTempId, TrayTempNo, TrayTempName, TrayTempCapacity, Remark
                FROM MES.TrayTemp 
                WHERE TrayTempNo = @TrayTempNo 
                AND [Status] = 'A' 
                AND CompanyId = @CompanyId
                AND NOT EXISTS (
                    SELECT 1 
                    FROM MES.Tray 
                    WHERE TrayNo = @TrayTempNo 
                    AND CompanyId = @CompanyId
                )
                ";

            var selectParams = new DynamicParameters();
            selectParams.Add("CompanyId", companyId);
            selectParams.Add("TrayTempNo", trayNo);

            var trayTempData = await connection.QueryAsync(selectSql, selectParams);
            var trayTempList = new List<dynamic>(trayTempData);

            if (trayTempList.Count == 0) return 0;

            // 批量插入到 Tray 表 (修正Lambda問題)
            var insertSql = @"
                INSERT INTO MES.Tray (CompanyId, BarcodeNo, TrayNo, TrayName, TrayCapacity, Remark, 
                                      UseTimes, Status, CreateDate, LastModifiedDate, 
                                      CreateBy, LastModifiedBy)                                
                VALUES (@CompanyId, @BarcodeNo, @TrayNo, @TrayName, @TrayCapacity, @Remark, 
                        @UseTimes, @Status, @CreateDate, @LastModifiedDate, 
                        @CreateBy, @LastModifiedBy)";

            for (int i = 0; i < trayTempList.Count; i++)
            {
                var item = trayTempList[i];
                var insertParams = new DynamicParameters();
                insertParams.Add("CompanyId", item.CompanyId);
                insertParams.Add("BarcodeNo", nullData);
                insertParams.Add("TrayNo", item.TrayTempNo);
                insertParams.Add("TrayName", item.TrayTempName);
                insertParams.Add("TrayCapacity", item.TrayTempCapacity);
                insertParams.Add("Remark", item.Remark ?? "");
                insertParams.Add("UseTimes", 0);
                insertParams.Add("Status", "A");
                insertParams.Add("CreateDate", DateTime.Now);
                insertParams.Add("LastModifiedDate", DateTime.Now);
                insertParams.Add("CreateBy", userId);
                insertParams.Add("LastModifiedBy", userId);

                var rowsInserted = await connection.ExecuteAsync(insertSql, insertParams);
                if (rowsInserted > 0) rowsAffected += rowsInserted;
            }

            return rowsAffected;
        }
        #endregion

        #region //GetLaserEngravedUserAsync //取得雷射刻印人員 (非同步版本) --Optimized 2025.03.06
        public async Task<string> GetLaserEngravedUserAsync(string Company, string LoginNo, string PassWord)
        {
            try
            {
                using (var sqlConnection = new SqlConnection(GetOptimizedConnectionString()))
                {
                    await sqlConnection.OpenAsync();

                    var sql = @"
                    SELECT a.RoleType, b.UserNo
                    FROM MES.LaserEngravedUser a
                    INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                    INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                    INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                    WHERE a.LoginNo = @LoginNo 
                      AND a.PassWord = @PassWord
                      AND d.CompanyNo = @Company";

                    var parameters = new DynamicParameters();
                    parameters.Add("LoginNo", LoginNo);
                    parameters.Add("PassWord", PassWord);
                    parameters.Add("Company", Company);

                    var result = await sqlConnection.QueryAsync(sql, parameters);

                    if (!result.Any())
                        throw new SystemException("此人員不具備雷刻機使用資格!");

                    #region //Response - 符合API文件格式
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response - 符合API文件格式
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

        #endregion
    }
}
