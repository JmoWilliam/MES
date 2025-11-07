using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Transactions;
using System.Web;
using System.Collections;

namespace QMSDA
{
    public class QcAnalysisToolsDA
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

        public QcAnalysisToolsDA()
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
        #region //GetQcMeasureDataQcItem -- 取得量測紀錄中的量測項目 -- WuTc -- 2024-06-24
        public string GetQcMeasureDataQcItem(int QcItemId, string StartDate, string EndDate, string MtlItemNo, string MtlItemName, string QcTypeNo, string CustomerMtlItemNo, string CustomerMtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.QcItemId, a.QcItemNo, a.QcItemName, b.QcItemDesc
                            , ISNULL(b.BallMark, '') AS BallMark
	                        , ISNULL(CONVERT(VARCHAR(18), b.DesignValue), '') AS DesignValue
	                        , ISNULL(CONVERT(VARCHAR(18), b.UpperTolerance), '') AS UpperTolerance
	                        , ISNULL(CONVERT(VARCHAR(18), b.LowerTolerance), '') AS LowerTolerance
                            , Qmm.MachineName + '(' + Qmm.MachineDesc + ')' QmmMachine
	                        FROM QMS.QcItem a
	                        INNER JOIN QMS.QcMeasureData b ON a.QcItemId = b.QcItemId	
	                        INNER JOIN MES.QcRecord c ON b.QcRecordId = c.QcRecordId
	                        INNER JOIN QMS.QcType d ON c.QcTypeId = d.QcTypeId
	                        OUTER APPLY ( SELECT ma.MoId, me.MtlItemNo, me.MtlItemName, wd.CompanyId
		                        FROM MES.ManufactureOrder ma
		                        INNER JOIN MES.WipOrder wd ON ma.WoId = wd.WoId
		                        INNER JOIN PDM.MtlItem me ON wd.MtlItemId = me.MtlItemId
		                        WHERE ma.MoId = c.MoId ) MtlItem
	                        OUTER APPLY	( SELECT ad.MachineNumber, ac.MachineNo, ac.MachineName, ac.MachineDesc 
		                        FROM MES.Machine ac
		                        INNER JOIN QMS.QmmDetail ad ON ac.MachineId = ad.MachineId
		                        WHERE b.QmmDetailId = ad.QmmDetailId ) Qmm
	                        WHERE MtlItem.CompanyId = @CurrentCompany";

                    dynamicParameters.Add("CurrentCompany", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND b.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND b.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND MtlItem.MtlItemNo LIKE '%' + @MtlItemNo + '%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND MtlItem.MtlItemName LIKE '%' + @MtlItemName + '%' ", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcTypeNo", @" AND d.QcTypeNo = @QcTypeNo ", QcTypeNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId ", QcItemId);
                    sql += @"    GROUP BY a.QcItemNo, a.QcItemName, b.QcItemDesc, DesignValue, UpperTolerance, LowerTolerance, Qmm.MachineName, Qmm.MachineDesc, a.QcItemId, b.BallMark
	                        ORDER BY a.QcItemNo";

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

        #region //GetQcMeasureDataforSPC -- 取得SPC 管制圖所需的量測資料 -- WuTc -- 2024-06-24
        public class ChartData
        {
            public int QcRecordId { get; set; }
            public DateTime date { get; set; }
            public string measureValue { get; set; }
            public string lettering { get; set; }
            public List<string> dateList { get; set; }
            public List<string> letteringList { get; set; }
            public List<string> measureValueList { get; set; }
            public int column { get; set; }
            public int row { get; set; }
        }
        public class Statistics
        {
            public string DynamicCenter { get; set; }
            public string DynamicUsl { get; set; }
            public string DynamicLsl { get; set; }
            public string dataCount { get; set; }
            public string n { get; set; }
            public string StdevS { get; set; }
            public string RCenter { get; set; }
            public string RUsl { get; set; }
            public string RLsl { get; set; }
            public string Ca { get; set; }
            public string Cp { get; set; }
            public string Cpk { get; set; }
            public string Pp { get; set; }
            public string Ppk { get; set; }
            public string GroupCount { get; set; }
        }
        public class Quartiles
        {
            public double Q1 { get; set; }
            public double Q2 { get; set; }
            public double Q3 { get; set; }
            public double QUpper { get; set; }
            public double QLower { get; set; }
        }
        public class JudgeStandards
        {
            public bool JudgeToF { get; set; }
            public string whitchJudge { get; set; }
            public List<string> Point { get; set; }
            public List<string> PointValue { get; set; }
            public List<string> PointDate { get; set; }
        }
        public string GetQcMeasureDataforSPC(List<int> judgeK, int QcItemId, string StartDate, string EndDate, string MtlItemNo, string MtlItemName, string QcTypeNo, string CustomerMtlItemNo, string CustomerMtlItemName, double DesignValue
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //sql取得量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT b.MoId, b.QcRecordId
	                        , a.CreateDate, a.QmdId, a.QcItemId, a.QcItemDesc, a.QmmDetailId
	                        , QcItem.QcItemNo, QcItem.QcItemName, QcItem.QcClassNo, QcItem.QcClassName
	                        , ISNULL(CONVERT(VARCHAR(18), a.DesignValue), '') AS DesignValue
	                        , ISNULL(CONVERT(VARCHAR(18), a.UpperTolerance), '') AS UpperTolerance
	                        , ISNULL(CONVERT(VARCHAR(18), a.LowerTolerance), '') AS LowerTolerance
	                        , ISNULL(a.ZAxis, '') ZAxis, a.CauseId, a.MeasureValue
	                        , ISNULL(a.BallMark, '') AS BallMark
                            , ISNULL(a.Unit, '') AS Unit
                            , Barcode.ItemSeq, a.Cavity, a.LotNumber
	                        , c.QcTypeNo, c.QcTypeName, b.InputType
	                        , MtlItem.MtlItemNo, MtlItem.MtlItemName
	                        , Qmm.MachineNo, Qmm.MachineNumber, Qmm.MachineName, Qmm.MachineDesc
	                        FROM QMS.QcMeasureData a
	                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
	                        INNER JOIN QMS.QcType c ON b.QcTypeId = c.QcTypeId
	                        OUTER APPLY ( SELECT ma.MoId, me.MtlItemNo, me.MtlItemName
		                        FROM MES.ManufactureOrder ma
		                        INNER JOIN MES.WipOrder wd ON ma.WoId = wd.WoId
		                        INNER JOIN PDM.MtlItem me ON wd.MtlItemId = me.MtlItemId
		                        WHERE ma.MoId = b.MoId ) MtlItem
	                        OUTER APPLY ( SELECT DISTINCT be.CurrentProdStatus, ba.ItemValue ,ba.ItemSeq
		                        FROM MES.BarcodeProcess bp
		                        INNER JOIN MES.Barcode be ON bp.BarcodeId = be.BarcodeId
		                        INNER JOIN MES.BarcodeAttribute ba ON be.BarcodeId = ba.BarcodeId
		                        WHERE be.MoId = MtlItem.MoId ) Barcode
	                        OUTER APPLY ( SELECT ae.QcItemNo, ae.QcItemName, ae.QcItemDesc, ag.QcClassNo, ag.QcClassName
		                        FROM QMS.QcItem ae
		                        INNER JOIN QMS.QcClass ag ON ae.QcClassId = ag.QcClassId
		                        WHERE a.QcItemId = ae.QcItemId ) QcItem
	                        OUTER APPLY	( SELECT ad.MachineNumber, ac.MachineNo, ac.MachineName, ac.MachineDesc 
		                        FROM MES.Machine ac
		                        INNER JOIN QMS.QmmDetail ad ON ac.MachineId = ad.MachineId
		                        WHERE a.QmmDetailId = ad.QmmDetailId ) Qmm
	                        WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND MtlItem.MtlItemNo LIKE '%' + @MtlItemNo + '%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND MtlItem.MtlItemName LIKE '%' + @MtlItemName + '%' ", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcTypeNo", @" AND c.QcTypeNo = @QcTypeNo ", QcTypeNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId ", QcItemId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DesignValue", @" AND a.DesignValue = @DesignValue ", DesignValue);
                    sql += @"   ORDER BY a.CreateDate, Barcode.ItemSeq, a.Cavity, a.LotNumber";

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    #endregion

                    List<double> measureValueList = new List<double>(); //統計涵數計算用
                    List<double> Statistics = new List<double>();
                    List<string> measureDateList = new List<string>();
                    List<string> letteringList = new List<string>();
                    List<ChartData> ChartDatas = new List<ChartData>();
                    List<ChartData> sortChartDatas = new List<ChartData>();
                    List<Statistics> statistics = new List<Statistics>();
                    double DesignUsl = 0;
                    double DesignLsl = 0;
                    int dataCount = 0;

                    //初步整理資料，將需要的資料抓出來，存入list
                    foreach (var item in result)
                    {
                        DateTime date = item.CreateDate;
                        string measureValue = item.MeasureValue;
                        string inputType = item.InputType;
                        int letteringSeq = item.ItemSeq ?? 0;
                        string cavity = item.Cavity;
                        string lotNumber = item.LotNumber;
                        string barcodeNo = item.BarcodeNo;
                        string letteringNo = "";
                        int qcRecordId = item.QcRecordId;
                        if (item.DesignValue == "") throw new SystemException("該筆資料設計值為空，無法產出SPC！");
                        DesignValue = Convert.ToDouble(item.DesignValue);
                        if (item.UpperTolerance == "") throw new SystemException("該筆資料上公差為空，無法產出SPC！");
                        DesignUsl = DesignValue + Convert.ToDouble(item.UpperTolerance);
                        if (item.LowerTolerance == "") throw new SystemException("該筆資料下公差為空，無法產出SPC！");
                        DesignLsl = DesignValue + Convert.ToDouble(item.LowerTolerance);
                        string IfSameData = date.ToString() + qcRecordId.ToString() + measureValue;
                        int isSame = 0;

                        //只抓數字，OK/NG不抓
                        bool isNumeric = double.TryParse(measureValue, out double num);
                        if (!isNumeric)
                        {
                            continue;
                        }

                        #region //letteringNo, 判斷刻字、穴號、批號、條碼
                        if (inputType == "Lettering")
                        {
                            letteringNo = letteringSeq.ToString(); //刻字
                        }
                        else if (inputType == "Cavity")
                        {
                            letteringNo = cavity; //穴號
                        }
                        else if (inputType == "LotNumber")
                        {
                            letteringNo = lotNumber; //批號
                        }
                        else if (inputType == "BarcodeNo")
                        {
                            letteringNo = barcodeNo.ToString(); //條碼
                        }
                        #endregion

                        ChartData chartData = new ChartData()
                        {
                            QcRecordId = qcRecordId,
                            date = date,
                            measureValue = measureValue,
                            lettering = letteringNo
                        };

                        foreach (var item2 in ChartDatas)
                        {
                            DateTime date2 = item2.date;
                            string measureValue2 = item2.measureValue;
                            int qcRecordId2 = item2.QcRecordId;
                            string CheckSameData = date2.ToString() + qcRecordId2.ToString() + measureValue2;

                            if (IfSameData == CheckSameData)
                            {
                                isSame = 1;
                                break;
                            }                            
                        }

                        if (isSame == 0)
                        {
                            letteringList.Add(letteringNo);
                            measureDateList.Add(date.ToShortDateString() + " (" + qcRecordId + ")");
                            measureValueList.Add(Convert.ToDouble(measureValue)); //統計涵數計算用
                            ChartDatas.Add(chartData);
                            dataCount++;
                        }
                    }
                    
                    #region //重新整理資料排列方式
                    int colIndex = 1;
                    int tableRowLength = letteringList.Distinct().Count();
                    int tableColLength = measureDateList.Distinct().Count();
                    List<string> measureDateListDIst = new List<string> (measureDateList.Distinct());
                    List<string> letteringListDIst = new List<string>(letteringList.Distinct());
                    ArrayList tableDatas = new ArrayList();
                    List<string> tableData = new List<string>();
                    tableData.Add("量測日期<br>(量測單號)");
                    //將同一天的數值放在同一個list裡
                    foreach (var item1 in measureDateListDIst)
                    {
                        string date = item1;
                        tableData.Add(item1);
                        List<string> letteringsubList = new List<string>();
                        List<string> measureValuesubList = new List<string>();
                        List<string> datesubList = new List<string>();
                        int qcRecordId = 0;
                        foreach (var item2 in ChartDatas)
                        {
                            int recordId = item2.QcRecordId;
                            //if 日期一樣，把value塞在同一個list裡，再塞進chartData
                            if (date == item2.date.ToShortDateString() + " (" + recordId.ToString() + ")")
                            {
                                qcRecordId = recordId;
                                string measureValue = item2.measureValue;
                                string letteringNo = item2.lettering;
                                datesubList.Add(item2.date.ToShortDateString() + " (" + qcRecordId.ToString() + ")");
                                measureValuesubList.Add(measureValue);
                                letteringsubList.Add(letteringNo);
                            }
                        }

                        ChartData chartData = new ChartData()
                        {
                            QcRecordId = qcRecordId,
                            date = Convert.ToDateTime(date.Split('(')[0]),
                            dateList = datesubList,
                            measureValueList = measureValuesubList,
                            letteringList = letteringsubList,
                            column = colIndex
                        };

                        sortChartDatas.Add(chartData);
                        colIndex += 1;
                    }
                    tableDatas.Add(tableData);

                    //將同一天的list，拆分成前端table好組的格式，並塞入表頭及表尾數據
                    //穴號為list第一碼，依序將不同日期的數據，塞入指定位置中
                    //table 的排列方式，同一個穴號+量測單號的數據，放在同一個list
                    int emptyCount = 0;
                    foreach (var item1 in letteringListDIst)
                    {
                        tableData = new List<string>();
                        int listCount = 0;
                        tableData.Add(item1); //先放入穴號/刻號在最左欄

                        foreach (var item2 in sortChartDatas)
                        {
                            int qcRecordId = 0;
                            string date = "";
                            string measureValue = "";
                            int loopCount = 0;
                            foreach (var lettering in item2.letteringList)
                            {
                                if (item1 == lettering)
                                {
                                    emptyCount = 0;
                                    qcRecordId = item2.QcRecordId;
                                    date = item2.date.ToShortDateString() + " (" + qcRecordId.ToString() + ")";

                                    foreach (var item3 in measureDateListDIst)
                                    {
                                        if (date == item3)
                                        {
                                            for (int i = tableData.Count -1; i < emptyCount; i++)
                                            {
                                                tableData.Add("");
                                            }
                                            measureValue = item2.measureValueList[loopCount];
                                            tableData.Add(measureValue);
                                            listCount += 1;
                                            break;
                                        }
                                        else
                                        {
                                            emptyCount += 1;
                                        }
                                    }
                                }
                                loopCount += 1;
                            }                            
                        }
                        for (int i = listCount + emptyCount; i < tableColLength; i++)
                        {
                            tableData.Add("");
                        }
                        tableDatas.Add(tableData);
                    }
                    #endregion

                    #region //統計涵數計算
                    #region //管制圖管制界限與中心值之因子系數表
                    int n = letteringList.Distinct().Count(); //統計因子系數表，只維護到n=25，如果n>25，要再新增系數
                    if (n > 25) throw new SystemException("統計因子n大於25，請檢視數據正確性，或洽系統開發室！");
                    int GroupCount = measureDateList.Distinct().Count();
                    double A = 0;
                    double A1 = 0;
                    double A2 = 0;
                    double C2 = 0;
                    double C4 = 0;
                    double B1 = 0;
                    double B2 = 0;
                    double B3 = 0;
                    double B4 = 0;
                    double d2 = 0;
                    double d3 = 0;
                    double D1 = 0;
                    double D2 = 0;
                    double D3 = 0;
                    double D4 = 0;
                    switch (n)
                    {
                        case 2:
                            A = 2.121; A1 = 3.760; A2 = 1.880;
                            C2 = 0.5642; C4 = 0.7979;
                            B1 = 0; B2 = 1.843; B3 = 0; B4 = 3.267;
                            d2 = 1.128; d3 = 0.853;
                            D1 = 0; D2 = 3.686; D3 = 0; D4 = 3.267;
                            break;
                        case 3:
                            A = 1.732; A1 = 2.394; A2 = 1.023;
                            C2 = 0.7236; C4 = 0.8862;
                            B1 = 0; B2 = 1.858; B3 = 0; B4 = 2.568;
                            d2 = 1.693; d3 = 0.888;
                            D1 = 0; D2 = 4.358; D3 = 0; D4 = 2.575;
                            break;
                        case 4:
                            A = 1.500; A1 = 1.880; A2 = 0.729;
                            C2 = 0.7979; C4 = 0.9213;
                            B1 = 0; B2 = 1.808; B3 = 0; B4 = 2.266;
                            d2 = 2.059; d3 = 0.880;
                            D1 = 0; D2 = 4.698; D3 = 0; D4 = 2.282;
                            break;
                        case 5:
                            A = 1.342; A1 = 1.596; A2 = 0.577;
                            C2 = 0.8407; C4 = 0.9400;
                            B1 = 0; B2 = 1.756; B3 = 0; B4 = 2.089;
                            d2 = 2.326; d3 = 0.864;
                            D1 = 0; D2 = 4.918; D3 = 0; D4 = 2.115;
                            break;
                        case 6:
                            A = 1.225; A1 = 1.410; A2 = 0.483;
                            C2 = 0.8686; C4 = 0.9515;
                            B1 = 0.026; B2 = 1.711; B3 = 0.030; B4 = 1.970;
                            d2 = 2.534; d3 = 0.848;
                            D1 = 0; D2 = 5.078; D3 = 0; D4 = 2.004;
                            break;
                        case 7:
                            A = 1.134; A1 = 1.277; A2 = 0.419;
                            C2 = 0.8882; C4 = 0.9594;
                            B1 = 0105; B2 = 1.672; B3 = 0.118; B4 = 1.882;
                            d2 = 2.704; d3 = 0.833;
                            D1 = 0.205; D2 = 5.203; D3 = 0.076; D4 = 1.924;
                            break;
                        case 8:
                            A = 1.061; A1 = 1.175; A2 = 0.373;
                            C2 = 0.9027; C4 = 0.9650;
                            B1 = 0.167; B2 = 1.638; B3 = 0.185; B4 = 1.815;
                            d2 = 2.847; d3 = 0.820;
                            D1 = 0.387; D2 = 5.307; D3 = 0.136; D4 = 1.864;
                            break;
                        case 9:
                            A = 1.000; A1 = 1.094; A2 = 0.337;
                            C2 = 0.9139; C4 = 0.9693;
                            B1 = 0.219; B2 = 1.609; B3 = 0.239; B4 = 1.761;
                            d2 = 2.970; d3 = 0.808;
                            D1 = 0.546; D2 = 5.394; D3 = 0.184; D4 = 1.816;
                            break;
                        case 10:
                            A = 0.949; A1 = 1.028; A2 = 0.308;
                            C2 = 0.9227; C4 = 0.9727;
                            B1 = 0.262; B2 = 1.584; B3 = 0.284; B4 = 1.716;
                            d2 = 3.078; d3 = 0.797;
                            D1 = 0.687; D2 = 5.469; D3 = 0.233; D4 = 1.777;
                            break;
                        case 11:
                            A = 0.905; A1 = 0.973; A2 = 0.285;
                            C2 = 0.9300; C4 = 0.9754;
                            B1 = 0.299; B2 = 1.561; B3 = 0.321; B4 = 1.679;
                            d2 = 3.173; d3 = 0.787;
                            D1 = 0.812; D2 = 5.534; D3 = 0.256; D4 = 1.744;
                            break;
                        case 12:
                            A = 0.866; A1 = 0.925; A2 = 0.266;
                            C2 = 0.9359; C4 = 0.9776;
                            B1 = 0.331; B2 = 1.541; B3 = 0.354; B4 = 1.646;
                            d2 = 3.258; d3 = 0.778;
                            D1 = 0.924; D2 = 5.592; D3 = 0.284; D4 = 1.716;
                            break;
                        case 13:
                            A = 0.832; A1 = 0.884; A2 = 0.249;
                            C2 = 0.9410; C4 = 0.9794;
                            B1 = 0.359; B2 = 1.523; B3 = 0.382; B4 = 1.618;
                            d2 = 3.336; d3 = 0.770;
                            D1 = 1.026; D2 = 5.636; D3 = 0.308; D4 = 1.692;
                            break;
                        case 14:
                            A = 0.802; A1 = 0.848; A2 = 0.235;
                            C2 = 0.9453; C4 = 0.9810;
                            B1 = 0.384; B2 = 1.507; B3 = 0.406; B4 = 1.594;
                            d2 = 3.407; d3 = 0.762;
                            D1 = 1.121; D2 = 5.693; D3 = 0.329; D4 = 1.671;
                            break;
                        case 15:
                            A = 0.775; A1 = 0.816; A2 = 0.223;
                            C2 = 0.9490; C4 = 0.9823;
                            B1 = 0.406; B2 = 1.492; B3 = 0.428; B4 = 1.572;
                            d2 = 3.472; d3 = 0.755;
                            D1 = 1.207; D2 = 5.737; D3 = 0.348; D4 = 1.652;
                            break;
                        case 16:
                            A = 0.750; A1 = 0.788; A2 = 0.212;
                            C2 = 0.9523; C4 = 0.9835;
                            B1 = 0.427; B2 = 1.478; B3 = 0.448; B4 = 1.552;
                            d2 = 3.532; d3 = 0.749;
                            D1 = 1.285; D2 = 5.779; D3 = 0.364; D4 = 1.636;
                            break;
                        case 17:
                            A = 0.728; A1 = 0.762; A2 = 0.203;
                            C2 = 0.9551; C4 = 0.9845;
                            B1 = 0.445; B2 = 1.465; B3 = 0.466; B4 = 1.534;
                            d2 = 3.588; d3 = 0.743;
                            D1 = 1.359; D2 = 5.817; D3 = 0.379; D4 = 1.621;
                            break;
                        case 18:
                            A = 0.707; A1 = 0.738; A2 = 0.194;
                            C2 = 0.9576; C4 = 0.9854;
                            B1 = 0.461; B2 = 1.454; B3 = 0.482; B4 = 1.518;
                            d2 = 3.640; d3 = 0.738;
                            D1 = 1.426; D2 = 5.854; D3 = 0.392; D4 = 1.606;
                            break;
                        case 19:
                            A = 0.688; A1 = 0.717; A2 = 0.187;
                            C2 = 0.9599; C4 = 0.9862;
                            B1 = 0.477; B2 = 1.443; B3 = 0.497; B4 = 1.503;
                            d2 = 3.689; d3 = 0.733;
                            D1 = 1.49; D2 = 5.888; D3 = 0.404; D4 = 1.596;
                            break;
                        case 20:
                            A = 0.671; A1 = 0.697; A2 = 0.180;
                            C2 = 0.9619; C4 = 0.9869;
                            B1 = 0.491; B2 = 1.433; B3 = 0.510; B4 = 1.490;
                            d2 = 3.735; d3 = 0.729;
                            D1 = 1.548; D2 = 5.922; D3 = 0.414; D4 = 1.586;
                            break;
                        case 21:
                            A = 0.655; A1 = 0.679; A2 = 0.173;
                            C2 = 0.9638; C4 = 0.9876;
                            B1 = 0.504; B2 = 1.424; B3 = 0.523; B4 = 1.477;
                            d2 = 3.778; d3 = 0.724;
                            D1 = 1.606; D2 = 5.950; D3 = 0.425; D4 = 1.575;
                            break;
                        case 22:
                            A = 0.640; A1 = 0.662; A2 = 0.167;
                            C2 = 0.9655; C4 = 0.9882;
                            B1 = 0.516; B2 = 1.415; B3 = 0.534; B4 = 1.466;
                            d2 = 3.819; d3 = 0.720;
                            D1 = 1.659; D2 = 5.979; D3 = 0.434; D4 = 1.586;
                            break;
                        case 23:
                            A = 0.626; A1 = 0.647; A2 = 0.162;
                            C2 = 0.9070; C4 = 0.9887;
                            B1 = 0.527; B2 = 1.427; B3 = 0.545; B4 = 1.455;
                            d2 = 3.885; d3 = 0.716;
                            D1 = 1.710; D2 = 6.006; D3 = 0.443; D4 = 1.557;
                            break;
                        case 24:
                            A = 0.612; A1 = 0.632; A2 = 0.157;
                            C2 = 0.9684; C4 = 0.9892;
                            B1 = 0.538; B2 = 1.399; B3 = 0.555; B4 = 1.445;
                            d2 = 3.895; d3 = 0.712;
                            D1 = 1.759; D2 = 6.031; D3 = 0.452; D4 = 1.548;
                            break;
                        case 25:
                            A = 0.600; A1 = 0.919; A2 = 0.153;
                            C2 = 0.9696; C4 = 0.9896;
                            B1 = 0.548; B2 = 1.392; B3 = 0.565; B4 = 1.435;
                            d2 = 3.931; d3 = 0.709;
                            D1 = 1.804; D2 = 6.058; D3 = 0.459; D4 = 1.541;
                            break;
                    }
                    #endregion

                    //Stdev.P 標準差
                    double MeasureAllValueAvg = measureValueList.Average();
                    double sumStdev = 0;
                    for (int i = 0; i < dataCount; i++)
                    {
                        sumStdev += (measureValueList[i] - MeasureAllValueAvg) * (measureValueList[i] - MeasureAllValueAvg);
                    }
                    double StdevS = Math.Sqrt(sumStdev / (dataCount - 1));
                    double StdevN = (DesignUsl - DesignLsl) / 12;

                    if (StdevS == 0) throw new SystemException("數據量不足，請重新選擇!"); ;

                    //計算每個日期的X Bar
                    List<string> tableCalculate = new List<string>(); //尾端的計算欄位，X Bar           
                    List<Quartiles> Quartiles = new List<Quartiles>();
                    tableCalculate.Add("X Bar");
                    foreach (var item1 in sortChartDatas)
                    {
                        double sumX = 0;
                        double XBar = 0;
                        int listCount = 0;
                        List<double> QuartileValue = new List<double>();

                        foreach (var item2 in item1.measureValueList)
                        {
                            QuartileValue.Add(Convert.ToDouble(item2));
                            sumX += Math.Round(Convert.ToDouble(item2), 4);
                            listCount += 1;
                        }
                        XBar = Math.Round(sumX / listCount, 4);
                        tableCalculate.Add(XBar.ToString());
                        Quartiles.Add(CalculateQuartiles(QuartileValue));
                    }
                    tableDatas.Add(tableCalculate);
                                        
                    //計算每個日期的R值
                    List<string> tableQuartiles = new List<string>();
                    tableQuartiles.Add("R");
                    double sumR = 0;
                    int countR = 0;
                    foreach (var item in Quartiles)
                    {
                        double max = item.QUpper;
                        double min = item.QLower;
                        double R = max - min;
                        sumR += R;
                        countR += 1;
                        tableQuartiles.Add(Math.Round(R, 5).ToString());
                    }
                    tableDatas.Add(tableQuartiles);

                    //管制界限：上限/下限/中心 R Chart
                    double RCenter = Math.Round(sumR / countR, 5);
                    double RUsl = Math.Round(D4 * RCenter, 5);
                    double RLsl = Math.Round(D3 * RCenter, 5);
                    //管制界限：上限/下限/中心 X Bar
                    double DynamicCenter = Math.Round(MeasureAllValueAvg, 4);
                    double DynamicUsl = Math.Round(DynamicCenter + (A2 * RCenter), 4);
                    double DynamicLsl = Math.Round(DynamicCenter - (A2 * RCenter), 4);

                    //計算製程能力指數：Ca、Cp、Cpk、Pp、Ppk
                    double Ca = Math.Round((DesignValue - DynamicCenter) / ((DesignUsl - DesignLsl) / 2), 4); //Ca = (M-X)/(T/2) M：產品中心位置（規格中心）X：群體的中心（平均值）T：規格寬度（規格上限-規格下限）
                    double Cp = 0; //Cp = T/(6σp)，T= 規格上限(USL) – 規格下限(LSL)、Cp=(USL-X)/(3σp)、Cp=(X-LSL)/(3σp) X為群體中心(平均值)
                    if (DesignUsl == 0) //CpL
                    {
                        Cp = Math.Round((MeasureAllValueAvg - DesignLsl) / (RCenter / d2 * 3), 4);
                    }
                    else if (DesignLsl == 0) //CpU
                    {
                        Cp = Math.Round((DesignUsl - MeasureAllValueAvg) / (RCenter / d2 * 3), 4);
                    }
                    else //CP
                    {
                        Cp = Math.Round((DesignUsl - DesignLsl) / (RCenter / d2 * 6), 4);
                    }
                    double Cpk = Math.Round(Cp * (1 - Math.Abs(Ca)), 4);
                    double Pp = Math.Round((DesignUsl - DesignLsl) / (StdevS * 6), 4);
                    double Ppk = Math.Round(Math.Min((DesignUsl - DynamicCenter) / (StdevS * 3), (DynamicCenter - DesignLsl) / (StdevS * 3)) * Pp, 4);

                    Statistics statistic = new Statistics
                    {
                        dataCount = dataCount.ToString(),
                        DynamicCenter = DynamicCenter.ToString(),
                        DynamicUsl = DynamicUsl.ToString(),
                        DynamicLsl = DynamicLsl.ToString(),
                        n = n.ToString(),
                        GroupCount = GroupCount.ToString(),
                        StdevS = StdevS.ToString(),
                        RCenter = RCenter.ToString(),
                        RUsl = RUsl.ToString(),
                        RLsl = RLsl.ToString(),
                        Ca = Ca.ToString(),
                        Cp = Cp.ToString(),
                        Cpk = Cpk.ToString(),
                        Pp = Pp.ToString(),
                        Ppk = Ppk.ToString()
                    };
                    statistics.Add(statistic);
                    #endregion

                    #region //超規判定及管制圖8種判定規則
                    #region //宣告
                    int One = judgeK[0];
                    int Two = judgeK[1];
                    int Three = judgeK[2];
                    int Four = judgeK[3];
                    int Five = judgeK[4];
                    int Six = judgeK[5];
                    int Seven = judgeK[6];
                    int Eight = judgeK[7];

                    List<JudgeStandards> judgeStandardXBars = new List<JudgeStandards>();
                    List<JudgeStandards> judgeStandardRs = new List<JudgeStandards>();
                    JudgeStandards judgeStandardXBar = new JudgeStandards();
                    JudgeStandards judgeStandardR = new JudgeStandards();
                    bool JudgeToF = false;
                    ArrayList tablefontColors = new ArrayList();
                    List<string> OnePointXBar = new List<string>();
                    List<string> OneValueXBar = new List<string>();
                    List<string> OneDateXBar = new List<string>();
                    List<string> TwoPointXBar = new List<string>();
                    List<string> TwoValueXBar = new List<string>();
                    List<string> TwoDateXBar = new List<string>();
                    List<string> ThreePointXBar = new List<string>();
                    List<string> ThreeValueXBar = new List<string>();
                    List<string> ThreeDateXBar = new List<string>();
                    List<string> FourPointXBar = new List<string>();
                    List<string> FourValueXBar = new List<string>();
                    List<string> FourDateXBar = new List<string>();
                    List<string> FivePointXBar = new List<string>();
                    List<string> FiveValueXBar = new List<string>();
                    List<string> FiveDateXBar = new List<string>();
                    List<string> SixPointXBar = new List<string>();
                    List<string> SixValueXBar = new List<string>();
                    List<string> SixDateXBar = new List<string>();
                    List<string> SevenPointXBar = new List<string>();
                    List<string> SevenValueXBar = new List<string>();
                    List<string> SevenDateXBar = new List<string>();
                    List<string> EightPointXBar = new List<string>();
                    List<string> EightValueXBar = new List<string>();
                    List<string> EightDateXBar = new List<string>();
                    List<string> OnePointR = new List<string>();
                    List<string> OneValueR = new List<string>();
                    List<string> OneDateR = new List<string>();
                    List<string> TwoPointR = new List<string>();
                    List<string> TwoValueR = new List<string>();
                    List<string> TwoDateR = new List<string>();
                    List<string> ThreePointR = new List<string>();
                    List<string> ThreeValueR = new List<string>();
                    List<string> ThreeDateR = new List<string>();
                    List<string> FourPointR = new List<string>();
                    List<string> FourValueR = new List<string>();
                    List<string> FourDateR = new List<string>();
                    List<string> FivePointR = new List<string>();
                    List<string> FiveValueR = new List<string>();
                    List<string> FiveDateR = new List<string>();
                    List<string> SixPointR = new List<string>();
                    List<string> SixValueR = new List<string>();
                    List<string> SixDateR = new List<string>();
                    List<string> SevenPointR = new List<string>();
                    List<string> SevenValueR = new List<string>();
                    List<string> SevenDateR = new List<string>();
                    List<string> EightPointR = new List<string>();
                    List<string> EightValueR = new List<string>();
                    List<string> EightDateR = new List<string>();
                    List<string> dateList = new List<string>();
                    int judgeCount = 0;
                    #endregion

                    foreach (List<string> item in tableDatas)
                    {
                        if (judgeCount == 0)
                        {
                            dateList = item;
                            judgeCount += 1;
                            continue;
                        }
                        List<string> tablefontColor = new List<string>();
                        string colorString = "";
                        int twoPlus = 0;
                        int twoMinus = 0;
                        int threePlus = 0;
                        int threeMinus = 0;
                        int fourPlus = 0;
                        int fourMinus = 0;
                        int fivePlus = 0;
                        int fiveMinus = 0;
                        int sixPlus = 0;
                        int sixMinus = 0;
                        int sevenPlus = 0;
                        int eightPlus = 0;
                        double LastThree = 0;
                        double LastFour = 0;
                        //量測數據
                        for (int i = 1; i < item.Count; i++)
                        {
                            if (item[i] != "")
                            {
                                double tableValue = Convert.ToDouble(item[i]);
                                if (judgeCount < tableDatas.Count - 2)
                                {
                                    if (tableValue < DesignLsl || tableValue > DesignUsl)
                                    {
                                        colorString = "#CE0000";
                                    }
                                    else
                                    {
                                        colorString = "#000000";
                                    }
                                }
                                else if (judgeCount == tableDatas.Count - 2) //X Bar
                                {
                                    if (tableValue < DynamicLsl || tableValue > DynamicUsl)
                                    {
                                        colorString = "##F75000";
                                    }
                                    else
                                    {
                                        colorString = "#000000";
                                    }
                                }
                                else if (judgeCount == tableDatas.Count - 1) //R
                                {
                                    if (tableValue < RLsl || tableValue > RUsl)
                                    {
                                        colorString = "##984B4B";
                                    }
                                    else
                                    {
                                        colorString = "#000000";
                                    }
                                }
                                tablefontColor.Add(colorString);
                            }
                            else
                            {
                                tablefontColor.Add("#000000");
                            }
                        }
                        tablefontColors.Add(tablefontColor);

                        #region //X Bar、R，8種判定規則
                        if (judgeCount >= tableDatas.Count - 2)
                        {
                            //A區為中心+-2σ到中心+-3σ之間
                            //B區為中心+-1σ到中心+-2σ之間
                            //C區為中心到中心+-1σ之間
                            //標準1：1點落在A區外
                            //標準2：連續九點落在同一側
                            //標準3：連續六點遞增或遞減
                            //標準4：14交替，連續十四點，相鄰點上下交錯
                            //標準5：2/3A，連續三點，有兩點落在同一側B區外(A區)
                            //標準6：4/5C，連續五點，有四點落在C區外
                            //標準7：15全C，連續十五點全在兩側C區內
                            //標準8：8缺C，連續八點落在兩側，但沒有任一點在C區
                            //A區: > DynamicCenter +- StdevS * 2;
                            //B區: > DynamicCenter +- StdevS * 1，< DynamicCenter +- StdevS * 2;
                            //C區: > DynamicCenter，< DynamicCenter +- StdevS * 1;
                            for (int i = 1; i < item.Count; i++)
                            {
                                if (item[i] != "")
                                {
                                    double tableValue = Convert.ToDouble(item[i]);
                                    string date = dateList[i].ToString();

                                    if (item[0] == "X Bar")
                                    {
                                        if (One != 0)
                                        {
                                            double StdevRate = Math.Round(StdevS * One, 5);
                                            //邏輯：數值落在A區，就放進LIST
                                            if (tableValue > (DynamicCenter + StdevRate) || tableValue < (DynamicCenter - StdevRate))
                                            {
                                                OnePointXBar.Add(i.ToString());
                                                OneDateXBar.Add(date);
                                                OneValueXBar.Add(tableValue.ToString());
                                            }
                                        }
                                        if (Two != 0)
                                        {
                                            //邏輯1：正與負各自計算，若數據跳到另一邊，則歸0重計
                                            //邏輯2：若plus or minus 小於 Two 則繼續反覆跳
                                            //邏輯3：若滿足 plus or minus 大於等於 Two 的條件，則另一邊不加入TwoJudgeXBar LIST
                                            if (tableValue > DynamicCenter)
                                            {
                                                if (twoMinus < Two)
                                                {
                                                    if (twoPlus == 0)
                                                    {
                                                        TwoPointXBar.Clear();
                                                        TwoValueXBar.Clear();
                                                        TwoDateXBar.Clear();
                                                    }
                                                    TwoPointXBar.Add(i.ToString());
                                                    TwoValueXBar.Add(tableValue.ToString());
                                                    TwoDateXBar.Add(date);
                                                    twoPlus += 1;
                                                    twoMinus = 0;
                                                }                                                    
                                            }
                                            else
                                            {
                                                if (twoPlus < Two)
                                                {
                                                    if (twoMinus == 0)
                                                    {
                                                        TwoPointXBar.Clear();
                                                        TwoValueXBar.Clear();
                                                        TwoDateXBar.Clear();
                                                    }
                                                    TwoPointXBar.Add(i.ToString());
                                                    TwoValueXBar.Add(tableValue.ToString());
                                                    TwoDateXBar.Add(date);
                                                    twoMinus += 1;
                                                    twoPlus = 0;
                                                }
                                            }
                                        }
                                        if (Three != 0)
                                        {
                                            //邏輯1：同Two的邏輯，但要與上一個數據對比
                                            if (tableValue >= LastThree)
                                            {
                                                if (threeMinus < Three)
                                                {
                                                    if (threePlus == 0)
                                                    {
                                                        ThreePointXBar.Clear();
                                                        ThreeValueXBar.Clear();
                                                        ThreeDateXBar.Clear();
                                                    }
                                                    ThreePointXBar.Add(i.ToString());
                                                    ThreeValueXBar.Add(tableValue.ToString());
                                                    ThreeDateXBar.Add(date);
                                                    threePlus += 1;
                                                    threeMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (threePlus < Three)
                                                {
                                                    if (threeMinus == 0)
                                                    {
                                                        ThreePointXBar.Clear();
                                                        ThreeValueXBar.Clear();
                                                        ThreeDateXBar.Clear();
                                                    }
                                                    ThreePointXBar.Add(i.ToString());
                                                    ThreeValueXBar.Add(tableValue.ToString());
                                                    ThreeDateXBar.Add(date);
                                                    threeMinus += 1;
                                                    threePlus = 0;
                                                }                                                
                                            }

                                            LastThree = tableValue;
                                        }
                                        if (Four != 0)
                                        {
                                            if (tableValue >= LastFour)
                                            {
                                                if (fourPlus > 0)
                                                {
                                                    fourPlus = 0;
                                                    FourPointXBar.Clear();
                                                    FourValueXBar.Clear();
                                                    FourDateXBar.Clear();
                                                }
                                                FourPointXBar.Add(i.ToString());
                                                FourValueXBar.Add(tableValue.ToString());
                                                FourDateXBar.Add(date);
                                                fourPlus += 1;
                                                fourMinus = 0;
                                            }
                                            else
                                            {
                                                if (fourMinus > 0)
                                                {
                                                    fourMinus = 0;
                                                    FourPointXBar.Clear();
                                                    FourValueXBar.Clear();
                                                    FourDateXBar.Clear();
                                                }
                                                FourPointXBar.Add(i.ToString());
                                                FourValueXBar.Add(tableValue.ToString());
                                                FourDateXBar.Add(date);
                                                fourMinus += 1;
                                                fourPlus = 0;
                                            }
                                            LastFour = tableValue;
                                        }
                                        if (Five != 0)
                                        {
                                            //邏輯1：連續 Five +1 中，有 Five 點超出StdevS * 2(A區)
                                            if (tableValue > (DynamicCenter + StdevS * 2))
                                            {
                                                if (fiveMinus < Five + 1)
                                                {
                                                    if (fivePlus == 0)
                                                    {
                                                        FivePointXBar.Clear();
                                                        FiveValueXBar.Clear();
                                                        FiveDateXBar.Clear();
                                                    }

                                                    FivePointXBar.Add(i.ToString());
                                                    FiveValueXBar.Add(tableValue.ToString());
                                                    FiveDateXBar.Add(date);
                                                    fivePlus += 1;
                                                    fiveMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (DynamicCenter - StdevS * 2))
                                            {
                                                if (fivePlus < Five + 1)
                                                {
                                                    if (fiveMinus == 0)
                                                    {
                                                        FivePointXBar.Clear();
                                                        FiveValueXBar.Clear();
                                                        FiveDateXBar.Clear();
                                                    }
                                                    FivePointXBar.Add(i.ToString());
                                                    FiveValueXBar.Add(tableValue.ToString());
                                                    FiveDateXBar.Add(date);
                                                    fiveMinus += 1;
                                                    fivePlus = 0;
                                                }
                                            }
                                        }
                                        if (Six != 0)
                                        {
                                            if (tableValue > (DynamicCenter + StdevS))
                                            {
                                                if (sixMinus < Six + 1)
                                                {
                                                    if (sixPlus == 0)
                                                    {
                                                        SixPointXBar.Clear();
                                                        SixValueXBar.Clear();
                                                        SixDateXBar.Clear();
                                                    }

                                                    SixPointXBar.Add(i.ToString());
                                                    SixValueXBar.Add(tableValue.ToString());
                                                    SixDateXBar.Add(date);
                                                    sixPlus += 1;
                                                    sixMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (DynamicCenter - StdevS))
                                            {
                                                if (sixPlus < Six + 1)
                                                {
                                                    if (sixMinus == 0)
                                                    {
                                                        SixPointXBar.Clear();
                                                        SixValueXBar.Clear();
                                                        SixDateXBar.Clear();
                                                    }
                                                    SixPointXBar.Add(i.ToString());
                                                    SixValueXBar.Add(tableValue.ToString());
                                                    SixDateXBar.Add(date);
                                                    sixMinus += 1;
                                                    sixPlus = 0;
                                                }
                                            }
                                        }
                                        if (Seven != 0)
                                        {
                                            if (tableValue < DynamicCenter + StdevS && tableValue > DynamicCenter - StdevS)
                                            {
                                                SevenPointXBar.Add(i.ToString());
                                                SevenValueXBar.Add(tableValue.ToString());
                                                SevenDateXBar.Add(date);
                                                sevenPlus += 1;
                                            }
                                            else
                                            {
                                                sevenPlus = 0;
                                                SevenPointXBar.Clear();
                                                SevenValueXBar.Clear();
                                                SevenDateXBar.Clear();
                                            }
                                        }
                                        if (Eight != 0)
                                        {
                                            if (tableValue > DynamicCenter + StdevS && tableValue < DynamicCenter - StdevS)
                                            {
                                                EightPointXBar.Add(i.ToString());
                                                EightValueXBar.Add(tableValue.ToString());
                                                EightDateXBar.Add(date);
                                                eightPlus += 1;
                                            }
                                            else
                                            {
                                                eightPlus = 0;
                                                EightPointXBar.Clear();
                                                EightValueXBar.Clear();
                                                EightDateXBar.Clear();
                                            }
                                        }
                                    }
                                    else if (item[0] == "R")
                                    {
                                        if (One != 0)
                                        {
                                            double StdevRate = Math.Round(StdevS * One, 5);
                                            //邏輯：數值落在A區，就放進LIST
                                            if (tableValue > (RCenter + StdevRate) || tableValue < (RCenter - StdevRate))                                                
                                            {
                                                OnePointR.Add(i.ToString());
                                                OneValueR.Add(tableValue.ToString());
                                                OneDateR.Add(date);
                                            }

                                        }
                                        if (Two != 0)
                                        {
                                            if (tableValue > RCenter)
                                            {
                                                if (twoMinus < Two)
                                                {
                                                    if (twoPlus == 0)
                                                    {
                                                        TwoPointR.Clear();
                                                        TwoValueR.Clear();
                                                        TwoDateR.Clear();
                                                    }
                                                    TwoPointR.Add(i.ToString());
                                                    TwoValueR.Add(tableValue.ToString());
                                                    TwoDateR.Add(date);
                                                    twoPlus += 1;
                                                    twoMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (twoPlus < Two)
                                                {
                                                    if (twoMinus == 0)
                                                    {
                                                        TwoPointR.Clear();
                                                        TwoValueR.Clear();
                                                        TwoDateR.Clear();
                                                    }
                                                    TwoPointR.Add(i.ToString());
                                                    TwoValueR.Add(tableValue.ToString());
                                                    TwoDateR.Add(date);
                                                    twoMinus += 1;
                                                    twoPlus = 0;
                                                }
                                            }
                                        }
                                        if (Three != 0)
                                        {
                                            if (tableValue >= LastThree)
                                            {
                                                if (threeMinus < Three)
                                                {
                                                    if (threePlus == 0)
                                                    {
                                                        ThreePointR.Clear();
                                                        ThreeValueR.Clear();
                                                        ThreeDateR.Clear();
                                                    }
                                                    ThreePointR.Add(i.ToString());
                                                    ThreeValueR.Add(tableValue.ToString());
                                                    ThreeDateR.Add(date);
                                                    threePlus += 1;
                                                    threeMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (threePlus < Three)
                                                {
                                                    if (threeMinus == 0)
                                                    {
                                                        ThreePointR.Clear();
                                                        ThreeValueR.Clear();
                                                        ThreeDateR.Clear();
                                                    }
                                                    ThreePointR.Add(i.ToString());
                                                    ThreeValueR.Add(tableValue.ToString());
                                                    ThreeDateR.Add(date);
                                                    threeMinus += 1;
                                                    threePlus = 0;
                                                }
                                            }

                                            LastThree = tableValue;
                                        }
                                        if (Four != 0)
                                        {
                                            if (tableValue >= LastFour)
                                            {
                                                if (fourPlus > 0)
                                                {
                                                    fourPlus = 0;
                                                    FourPointR.Clear();
                                                    FourValueR.Clear();
                                                    FourDateR.Clear();
                                                }
                                                FourPointR.Add(i.ToString());
                                                FourValueR.Add(tableValue.ToString());
                                                FourDateR.Add(date);
                                                fourPlus += 1;
                                                fourMinus = 0;
                                            }
                                            else
                                            {
                                                if (fourMinus > 0)
                                                {
                                                    fourMinus = 0;
                                                    FourPointR.Clear();
                                                    FourValueR.Clear();
                                                    FourDateR.Clear();
                                                }
                                                FourPointR.Add(i.ToString());
                                                FourValueR.Add(tableValue.ToString());
                                                FourDateR.Add(date);
                                                fourMinus += 1;
                                                fourPlus = 0;
                                            }
                                            LastFour = tableValue;
                                        }
                                        if (Five != 0)
                                        {
                                            if (tableValue > (RCenter + StdevS * 2))
                                            {
                                                if (fiveMinus < Five + 1)
                                                {
                                                    if (fivePlus == 0)
                                                    {
                                                        FivePointR.Clear();
                                                        FiveValueR.Clear();
                                                        FiveDateR.Clear();
                                                    }

                                                    FivePointR.Add(i.ToString());
                                                    FiveValueR.Add(tableValue.ToString());
                                                    FiveDateR.Add(date);
                                                    fivePlus += 1;
                                                    fiveMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (RCenter - StdevS * 2))
                                            {
                                                if (fivePlus < Five + 1)
                                                {
                                                    if (fiveMinus == 0)
                                                    {
                                                        FivePointR.Clear();
                                                        FiveValueR.Clear();
                                                        FiveDateR.Clear();
                                                    }
                                                    FivePointR.Add(i.ToString());
                                                    FiveValueR.Add(tableValue.ToString());
                                                    FiveDateR.Add(date);
                                                    fiveMinus += 1;
                                                    fivePlus = 0;
                                                }
                                            }
                                        }
                                        if (Six != 0)
                                        {
                                            if (tableValue > (RCenter + StdevS))
                                            {
                                                if (sixMinus < Six + 1)
                                                {
                                                    if (sixPlus == 0)
                                                    {
                                                        SixPointR.Clear();
                                                        SixValueR.Clear();
                                                        SixDateR.Clear();
                                                    }

                                                    SixPointR.Add(i.ToString());
                                                    SixValueR.Add(tableValue.ToString());
                                                    SixDateR.Add(date);
                                                    sixPlus += 1;
                                                    sixMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (RCenter - StdevS))
                                            {
                                                if (sixPlus < Six + 1)
                                                {
                                                    if (sixMinus == 0)
                                                    {
                                                        SixPointR.Clear();
                                                        SixValueR.Clear();
                                                        SixDateR.Clear();
                                                    }
                                                    SixPointR.Add(i.ToString());
                                                    SixValueR.Add(tableValue.ToString());
                                                    SixDateR.Add(date);
                                                    sixMinus += 1;
                                                    sixPlus = 0;
                                                }
                                            }
                                        }
                                        if (Seven != 0)
                                        {
                                            if (tableValue < RCenter + StdevS && tableValue > RCenter - StdevS)
                                            {
                                                SevenPointR.Add(i.ToString());
                                                SevenValueR.Add(tableValue.ToString());
                                                SevenDateR.Add(date);
                                                sevenPlus += 1;
                                            }
                                            else
                                            {
                                                sevenPlus = 0;
                                                SevenPointR.Clear();
                                                SevenValueR.Clear();
                                                SevenDateR.Clear();
                                            }
                                        }
                                        if (Eight != 0)
                                        {
                                            if (tableValue > RCenter + StdevS && tableValue < RCenter - StdevS)
                                            {
                                                EightPointR.Add(i.ToString());
                                                EightValueR.Add(tableValue.ToString());
                                                EightDateR.Add(date);
                                                eightPlus += 1;
                                            }
                                            else
                                            {
                                                eightPlus = 0;
                                                EightPointR.Clear();
                                                EightValueR.Clear();
                                                EightDateR.Clear();
                                            }
                                        }
                                    }
                                }
                            }
                        }                        
                        #endregion

                        judgeCount += 1;
                    }

                    #region //將判定結果放到list裡
                    //XBar
                    if (OnePointXBar.Count >= 1)
                    {
                        JudgeToF = true;
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 A",
                        Point = OnePointXBar,
                        PointDate = OneDateXBar,
                        PointValue = OneValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Two != 0)
                    {
                        if (TwoPointXBar.Count >= Two)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }                    
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 B",
                        Point = TwoPointXBar,
                        PointDate = TwoDateXBar,
                        PointValue = TwoValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Three != 0)
                    {
                        if (ThreePointXBar.Count >= Three)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }                    
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 C",
                        Point = ThreePointXBar,
                        PointDate = ThreeDateXBar,
                        PointValue = ThreeValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Four != 0)
                    {
                        if (FourPointXBar.Count >= Four)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }                    
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 D",
                        Point = FourPointXBar,
                        PointDate = FourDateXBar,
                        PointValue = FourValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Five != 0)
                    {
                        if (FivePointXBar.Count >= Five)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }                    
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 E",
                        Point = FivePointXBar,
                        PointDate = FiveDateXBar,
                        PointValue = FiveValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Six != 0)
                    {
                        if (SixPointXBar.Count >= Six)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 F",
                        Point = SixPointXBar,
                        PointDate = SixDateXBar,
                        PointValue = SixValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Seven != 0)
                    {
                        if (SevenPointXBar.Count >= Seven)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }                    
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 G",
                        Point = SevenPointXBar,
                        PointDate = SevenDateXBar,
                        PointValue = SevenValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Eight != 0)
                    {
                        if (EightPointXBar.Count >= Eight)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }                    
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 H",
                        Point = EightPointXBar,
                        PointDate = EightDateXBar,
                        PointValue = EightValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);

                    //R
                    if (OnePointR.Count >= 1)
                    {
                        JudgeToF = true;
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 A",
                        Point = OnePointR,
                        PointDate = OneDateR,
                        PointValue = OneValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Two != 0)
                    {
                        if (TwoPointR.Count >= Two)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 B",
                        Point = TwoPointR,
                        PointDate = TwoDateR,
                        PointValue = TwoValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Three != 0)
                    {
                        if (ThreePointR.Count >= Three)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 C",
                        Point = ThreePointR,
                        PointDate = ThreeDateR,
                        PointValue = ThreeValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Four != 0)
                    {
                        if (FourPointR.Count >= Four)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 D",
                        Point = FourPointR,
                        PointDate = FourDateR,
                        PointValue = FourValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Five != 0)
                    {
                        if (FivePointR.Count >= Five)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 E",
                        Point = FivePointR,
                        PointDate = FiveDateR,
                        PointValue = FiveValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Six != 0)
                    {
                        if (SixPointR.Count >= Six)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 F",
                        Point = SixPointR,
                        PointDate = SixDateR,
                        PointValue = SixValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Seven != 0)
                    {
                        if (SevenPointR.Count >= Seven)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 G",
                        Point = SevenPointR,
                        PointDate = SevenDateR,
                        PointValue = SevenValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Eight != 0)
                    {
                        if (EightPointR.Count >= Eight)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 H",
                        Point = EightPointR,
                        PointDate = EightDateR,
                        PointValue = EightValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    #endregion

                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = new { result, Quartiles, statistics, tableDatas, tablefontColors, judgeStandardXBars, judgeStandardRs }
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

        #region //GetQcMeasureDataforSPCExcel -- 取得SPC 資料匯出至excel  -- WuTc -- 2024-06-24
        public string GetQcMeasureDataforSPCExcel(List<int> judgeK, int QcItemId, string StartDate, string EndDate, string MtlItemNo, string MtlItemName, string QcTypeNo, string CustomerMtlItemNo, string CustomerMtlItemName, double DesignValue
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //sql取得量測數據
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT b.MoId, b.QcRecordId
	                        , a.CreateDate, a.QmdId, a.QcItemId, a.QcItemDesc, a.QmmDetailId
	                        , QcItem.QcItemNo, QcItem.QcItemName, QcItem.QcClassNo, QcItem.QcClassName
	                        , ISNULL(CONVERT(VARCHAR(18), a.DesignValue), '') AS DesignValue
	                        , ISNULL(CONVERT(VARCHAR(18), a.UpperTolerance), '') AS UpperTolerance
	                        , ISNULL(CONVERT(VARCHAR(18), a.LowerTolerance), '') AS LowerTolerance
	                        , ISNULL(a.ZAxis, '') ZAxis, a.CauseId, a.MeasureValue
	                        , ISNULL(a.BallMark, '') AS BallMark
                            , ISNULL(a.Unit, '') AS Unit
                            , Barcode.ItemSeq, a.Cavity, a.LotNumber
	                        , c.QcTypeNo, c.QcTypeName, b.InputType
	                        , MtlItem.MtlItemNo, MtlItem.MtlItemName
	                        , Qmm.MachineNo, Qmm.MachineNumber, Qmm.MachineName, Qmm.MachineDesc
	                        FROM QMS.QcMeasureData a
	                        INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
	                        INNER JOIN QMS.QcType c ON b.QcTypeId = c.QcTypeId
	                        OUTER APPLY ( SELECT ma.MoId, me.MtlItemNo, me.MtlItemName
		                        FROM MES.ManufactureOrder ma
		                        INNER JOIN MES.WipOrder wd ON ma.WoId = wd.WoId
		                        INNER JOIN PDM.MtlItem me ON wd.MtlItemId = me.MtlItemId
		                        WHERE ma.MoId = b.MoId ) MtlItem
	                        OUTER APPLY ( SELECT DISTINCT be.CurrentProdStatus, ba.ItemValue ,ba.ItemSeq
		                        FROM MES.BarcodeProcess bp
		                        INNER JOIN MES.Barcode be ON bp.BarcodeId = be.BarcodeId
		                        INNER JOIN MES.BarcodeAttribute ba ON be.BarcodeId = ba.BarcodeId
		                        WHERE be.MoId = MtlItem.MoId ) Barcode
	                        OUTER APPLY ( SELECT ae.QcItemNo, ae.QcItemName, ae.QcItemDesc, ag.QcClassNo, ag.QcClassName
		                        FROM QMS.QcItem ae
		                        INNER JOIN QMS.QcClass ag ON ae.QcClassId = ag.QcClassId
		                        WHERE a.QcItemId = ae.QcItemId ) QcItem
	                        OUTER APPLY	( SELECT ad.MachineNumber, ac.MachineNo, ac.MachineName, ac.MachineDesc 
		                        FROM MES.Machine ac
		                        INNER JOIN QMS.QmmDetail ad ON ac.MachineId = ad.MachineId
		                        WHERE a.QmmDetailId = ad.QmmDetailId ) Qmm
	                        WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND MtlItem.MtlItemNo LIKE '%' + @MtlItemNo + '%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND MtlItem.MtlItemName LIKE '%' + @MtlItemName + '%' ", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcTypeNo", @" AND c.QcTypeNo = @QcTypeNo ", QcTypeNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId ", QcItemId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DesignValue", @" AND a.DesignValue = @DesignValue ", DesignValue);
                    sql += @"   ORDER BY a.CreateDate, Barcode.ItemSeq, a.Cavity, a.LotNumber";

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    #endregion

                    List<double> measureValueList = new List<double>(); //統計涵數計算用
                    List<double> Statistics = new List<double>();
                    List<string> measureDateList = new List<string>();
                    List<string> letteringList = new List<string>();
                    List<ChartData> ChartDatas = new List<ChartData>();
                    List<ChartData> sortChartDatas = new List<ChartData>();
                    List<Statistics> statistics = new List<Statistics>();
                    double DesignUsl = 0;
                    double DesignLsl = 0;
                    int dataCount = 0;

                    //初步整理資料，將需要的資料抓出來，存入list
                    foreach (var item in result)
                    {
                        DateTime date = item.CreateDate;
                        string measureValue = item.MeasureValue;
                        string inputType = item.InputType;
                        int letteringSeq = item.ItemSeq ?? 0;
                        string cavity = item.Cavity;
                        string lotNumber = item.LotNumber;
                        string barcodeNo = item.BarcodeNo;
                        string letteringNo = "";
                        int qcRecordId = item.QcRecordId;
                        DesignValue = Convert.ToDouble(item.DesignValue);
                        DesignUsl = DesignValue + Convert.ToDouble(item.UpperTolerance);
                        DesignLsl = DesignValue + Convert.ToDouble(item.LowerTolerance);
                        string IfSameData = date.ToString() + qcRecordId.ToString() + measureValue;
                        int isSame = 0;

                        //只抓數字，OK/NG不抓
                        bool isNumeric = double.TryParse(measureValue, out double num);
                        if (!isNumeric)
                        {
                            continue;
                        }

                        #region //letteringNo, 判斷刻字、穴號、批號、條碼
                        if (inputType == "Lettering")
                        {
                            letteringNo = letteringSeq.ToString(); //刻字
                        }
                        else if (inputType == "Cavity")
                        {
                            letteringNo = cavity; //穴號
                        }
                        else if (inputType == "LotNumber")
                        {
                            letteringNo = lotNumber; //批號
                        }
                        else if (inputType == "BarcodeNo")
                        {
                            letteringNo = barcodeNo.ToString(); //條碼
                        }
                        #endregion

                        ChartData chartData = new ChartData()
                        {
                            QcRecordId = qcRecordId,
                            date = date,
                            measureValue = measureValue,
                            lettering = letteringNo
                        };

                        foreach (var item2 in ChartDatas)
                        {
                            DateTime date2 = item2.date;
                            string measureValue2 = item2.measureValue;
                            int qcRecordId2 = item2.QcRecordId;
                            string CheckSameData = date2.ToString() + qcRecordId2.ToString() + measureValue2;

                            if (IfSameData == CheckSameData)
                            {
                                isSame = 1;
                                break;
                            }
                        }

                        if (isSame == 0)
                        {
                            letteringList.Add(letteringNo);
                            measureDateList.Add(date.ToShortDateString() + " (" + qcRecordId + ")");
                            measureValueList.Add(Convert.ToDouble(measureValue)); //統計涵數計算用
                            ChartDatas.Add(chartData);
                            dataCount++;
                        }
                    }

                    #region //重新整理資料排列方式
                    int colIndex = 1;
                    int tableRowLength = letteringList.Distinct().Count();
                    int tableColLength = measureDateList.Distinct().Count();
                    List<string> measureDateListDIst = new List<string>(measureDateList.Distinct());
                    List<string> letteringListDIst = new List<string>(letteringList.Distinct());
                    ArrayList tableDatas = new ArrayList();
                    List<string> tableData = new List<string>();
                    tableData.Add("量測日期<br>(量測單號)");
                    //將同一天的數值放在同一個list裡
                    foreach (var item1 in measureDateListDIst)
                    {
                        string date = item1;
                        tableData.Add(item1);
                        List<string> letteringsubList = new List<string>();
                        List<string> measureValuesubList = new List<string>();
                        List<string> datesubList = new List<string>();
                        int qcRecordId = 0;
                        foreach (var item2 in ChartDatas)
                        {
                            int recordId = item2.QcRecordId;
                            //if 日期一樣，把value塞在同一個list裡，再塞進chartData
                            if (date == item2.date.ToShortDateString() + " (" + recordId.ToString() + ")")
                            {
                                qcRecordId = recordId;
                                string measureValue = item2.measureValue;
                                string letteringNo = item2.lettering;
                                datesubList.Add(item2.date.ToShortDateString() + " (" + qcRecordId.ToString() + ")");
                                measureValuesubList.Add(measureValue);
                                letteringsubList.Add(letteringNo);
                            }
                        }

                        ChartData chartData = new ChartData()
                        {
                            QcRecordId = qcRecordId,
                            date = Convert.ToDateTime(date.Split('(')[0]),
                            dateList = datesubList,
                            measureValueList = measureValuesubList,
                            letteringList = letteringsubList,
                            column = colIndex
                        };

                        sortChartDatas.Add(chartData);
                        colIndex += 1;
                    }
                    tableDatas.Add(tableData);

                    //將同一天的list，拆分成前端table好組的格式，並塞入表頭及表尾數據
                    //穴號為list第一碼，依序將不同日期的數據，塞入指定位置中
                    //table 的排列方式，同一個穴號+量測單號的數據，放在同一個list
                    int emptyCount = 0;
                    foreach (var item1 in letteringListDIst)
                    {
                        tableData = new List<string>();
                        int listCount = 0;
                        tableData.Add(item1); //先放入穴號/刻號在最左欄

                        foreach (var item2 in sortChartDatas)
                        {
                            int qcRecordId = 0;
                            string date = "";
                            string measureValue = "";
                            int loopCount = 0;
                            foreach (var lettering in item2.letteringList)
                            {
                                if (item1 == lettering)
                                {
                                    emptyCount = 0;
                                    qcRecordId = item2.QcRecordId;
                                    date = item2.date.ToShortDateString() + " (" + qcRecordId.ToString() + ")";

                                    foreach (var item3 in measureDateListDIst)
                                    {
                                        if (date == item3)
                                        {
                                            for (int i = tableData.Count - 1; i < emptyCount; i++)
                                            {
                                                tableData.Add("");
                                            }
                                            measureValue = item2.measureValueList[loopCount];
                                            tableData.Add(measureValue);
                                            listCount += 1;
                                            break;
                                        }
                                        else
                                        {
                                            emptyCount += 1;
                                        }
                                    }
                                }
                                loopCount += 1;
                            }
                        }
                        for (int i = listCount + emptyCount; i < tableColLength; i++)
                        {
                            tableData.Add("");
                        }
                        tableDatas.Add(tableData);
                    }
                    #endregion

                    #region //統計涵數計算
                    #region //管制圖管制界限與中心值之因子系數表
                    int n = letteringList.Distinct().Count(); //統計因子系數表，只維護到n=25，如果n>25，要再新增系數
                    if (n > 25) throw new SystemException("統計因子n大於25，請檢視數據正確性，或洽系統開發室！");
                    int GroupCount = measureDateList.Distinct().Count();
                    double A = 0;
                    double A1 = 0;
                    double A2 = 0;
                    double C2 = 0;
                    double C4 = 0;
                    double B1 = 0;
                    double B2 = 0;
                    double B3 = 0;
                    double B4 = 0;
                    double d2 = 0;
                    double d3 = 0;
                    double D1 = 0;
                    double D2 = 0;
                    double D3 = 0;
                    double D4 = 0;
                    switch (n)
                    {
                        case 2:
                            A = 2.121; A1 = 3.760; A2 = 1.880;
                            C2 = 0.5642; C4 = 0.7979;
                            B1 = 0; B2 = 1.843; B3 = 0; B4 = 3.267;
                            d2 = 1.128; d3 = 0.853;
                            D1 = 0; D2 = 3.686; D3 = 0; D4 = 3.267;
                            break;
                        case 3:
                            A = 1.732; A1 = 2.394; A2 = 1.023;
                            C2 = 0.7236; C4 = 0.8862;
                            B1 = 0; B2 = 1.858; B3 = 0; B4 = 2.568;
                            d2 = 1.693; d3 = 0.888;
                            D1 = 0; D2 = 4.358; D3 = 0; D4 = 2.575;
                            break;
                        case 4:
                            A = 1.500; A1 = 1.880; A2 = 0.729;
                            C2 = 0.7979; C4 = 0.9213;
                            B1 = 0; B2 = 1.808; B3 = 0; B4 = 2.266;
                            d2 = 2.059; d3 = 0.880;
                            D1 = 0; D2 = 4.698; D3 = 0; D4 = 2.282;
                            break;
                        case 5:
                            A = 1.342; A1 = 1.596; A2 = 0.577;
                            C2 = 0.8407; C4 = 0.9400;
                            B1 = 0; B2 = 1.756; B3 = 0; B4 = 2.089;
                            d2 = 2.326; d3 = 0.864;
                            D1 = 0; D2 = 4.918; D3 = 0; D4 = 2.115;
                            break;
                        case 6:
                            A = 1.225; A1 = 1.410; A2 = 0.483;
                            C2 = 0.8686; C4 = 0.9515;
                            B1 = 0.026; B2 = 1.711; B3 = 0.030; B4 = 1.970;
                            d2 = 2.534; d3 = 0.848;
                            D1 = 0; D2 = 5.078; D3 = 0; D4 = 2.004;
                            break;
                        case 7:
                            A = 1.134; A1 = 1.277; A2 = 0.419;
                            C2 = 0.8882; C4 = 0.9594;
                            B1 = 0105; B2 = 1.672; B3 = 0.118; B4 = 1.882;
                            d2 = 2.704; d3 = 0.833;
                            D1 = 0.205; D2 = 5.203; D3 = 0.076; D4 = 1.924;
                            break;
                        case 8:
                            A = 1.061; A1 = 1.175; A2 = 0.373;
                            C2 = 0.9027; C4 = 0.9650;
                            B1 = 0.167; B2 = 1.638; B3 = 0.185; B4 = 1.815;
                            d2 = 2.847; d3 = 0.820;
                            D1 = 0.387; D2 = 5.307; D3 = 0.136; D4 = 1.864;
                            break;
                        case 9:
                            A = 1.000; A1 = 1.094; A2 = 0.337;
                            C2 = 0.9139; C4 = 0.9693;
                            B1 = 0.219; B2 = 1.609; B3 = 0.239; B4 = 1.761;
                            d2 = 2.970; d3 = 0.808;
                            D1 = 0.546; D2 = 5.394; D3 = 0.184; D4 = 1.816;
                            break;
                        case 10:
                            A = 0.949; A1 = 1.028; A2 = 0.308;
                            C2 = 0.9227; C4 = 0.9727;
                            B1 = 0.262; B2 = 1.584; B3 = 0.284; B4 = 1.716;
                            d2 = 3.078; d3 = 0.797;
                            D1 = 0.687; D2 = 5.469; D3 = 0.233; D4 = 1.777;
                            break;
                        case 11:
                            A = 0.905; A1 = 0.973; A2 = 0.285;
                            C2 = 0.9300; C4 = 0.9754;
                            B1 = 0.299; B2 = 1.561; B3 = 0.321; B4 = 1.679;
                            d2 = 3.173; d3 = 0.787;
                            D1 = 0.812; D2 = 5.534; D3 = 0.256; D4 = 1.744;
                            break;
                        case 12:
                            A = 0.866; A1 = 0.925; A2 = 0.266;
                            C2 = 0.9359; C4 = 0.9776;
                            B1 = 0.331; B2 = 1.541; B3 = 0.354; B4 = 1.646;
                            d2 = 3.258; d3 = 0.778;
                            D1 = 0.924; D2 = 5.592; D3 = 0.284; D4 = 1.716;
                            break;
                        case 13:
                            A = 0.832; A1 = 0.884; A2 = 0.249;
                            C2 = 0.9410; C4 = 0.9794;
                            B1 = 0.359; B2 = 1.523; B3 = 0.382; B4 = 1.618;
                            d2 = 3.336; d3 = 0.770;
                            D1 = 1.026; D2 = 5.636; D3 = 0.308; D4 = 1.692;
                            break;
                        case 14:
                            A = 0.802; A1 = 0.848; A2 = 0.235;
                            C2 = 0.9453; C4 = 0.9810;
                            B1 = 0.384; B2 = 1.507; B3 = 0.406; B4 = 1.594;
                            d2 = 3.407; d3 = 0.762;
                            D1 = 1.121; D2 = 5.693; D3 = 0.329; D4 = 1.671;
                            break;
                        case 15:
                            A = 0.775; A1 = 0.816; A2 = 0.223;
                            C2 = 0.9490; C4 = 0.9823;
                            B1 = 0.406; B2 = 1.492; B3 = 0.428; B4 = 1.572;
                            d2 = 3.472; d3 = 0.755;
                            D1 = 1.207; D2 = 5.737; D3 = 0.348; D4 = 1.652;
                            break;
                        case 16:
                            A = 0.750; A1 = 0.788; A2 = 0.212;
                            C2 = 0.9523; C4 = 0.9835;
                            B1 = 0.427; B2 = 1.478; B3 = 0.448; B4 = 1.552;
                            d2 = 3.532; d3 = 0.749;
                            D1 = 1.285; D2 = 5.779; D3 = 0.364; D4 = 1.636;
                            break;
                        case 17:
                            A = 0.728; A1 = 0.762; A2 = 0.203;
                            C2 = 0.9551; C4 = 0.9845;
                            B1 = 0.445; B2 = 1.465; B3 = 0.466; B4 = 1.534;
                            d2 = 3.588; d3 = 0.743;
                            D1 = 1.359; D2 = 5.817; D3 = 0.379; D4 = 1.621;
                            break;
                        case 18:
                            A = 0.707; A1 = 0.738; A2 = 0.194;
                            C2 = 0.9576; C4 = 0.9854;
                            B1 = 0.461; B2 = 1.454; B3 = 0.482; B4 = 1.518;
                            d2 = 3.640; d3 = 0.738;
                            D1 = 1.426; D2 = 5.854; D3 = 0.392; D4 = 1.606;
                            break;
                        case 19:
                            A = 0.688; A1 = 0.717; A2 = 0.187;
                            C2 = 0.9599; C4 = 0.9862;
                            B1 = 0.477; B2 = 1.443; B3 = 0.497; B4 = 1.503;
                            d2 = 3.689; d3 = 0.733;
                            D1 = 1.49; D2 = 5.888; D3 = 0.404; D4 = 1.596;
                            break;
                        case 20:
                            A = 0.671; A1 = 0.697; A2 = 0.180;
                            C2 = 0.9619; C4 = 0.9869;
                            B1 = 0.491; B2 = 1.433; B3 = 0.510; B4 = 1.490;
                            d2 = 3.735; d3 = 0.729;
                            D1 = 1.548; D2 = 5.922; D3 = 0.414; D4 = 1.586;
                            break;
                        case 21:
                            A = 0.655; A1 = 0.679; A2 = 0.173;
                            C2 = 0.9638; C4 = 0.9876;
                            B1 = 0.504; B2 = 1.424; B3 = 0.523; B4 = 1.477;
                            d2 = 3.778; d3 = 0.724;
                            D1 = 1.606; D2 = 5.950; D3 = 0.425; D4 = 1.575;
                            break;
                        case 22:
                            A = 0.640; A1 = 0.662; A2 = 0.167;
                            C2 = 0.9655; C4 = 0.9882;
                            B1 = 0.516; B2 = 1.415; B3 = 0.534; B4 = 1.466;
                            d2 = 3.819; d3 = 0.720;
                            D1 = 1.659; D2 = 5.979; D3 = 0.434; D4 = 1.586;
                            break;
                        case 23:
                            A = 0.626; A1 = 0.647; A2 = 0.162;
                            C2 = 0.9070; C4 = 0.9887;
                            B1 = 0.527; B2 = 1.427; B3 = 0.545; B4 = 1.455;
                            d2 = 3.885; d3 = 0.716;
                            D1 = 1.710; D2 = 6.006; D3 = 0.443; D4 = 1.557;
                            break;
                        case 24:
                            A = 0.612; A1 = 0.632; A2 = 0.157;
                            C2 = 0.9684; C4 = 0.9892;
                            B1 = 0.538; B2 = 1.399; B3 = 0.555; B4 = 1.445;
                            d2 = 3.895; d3 = 0.712;
                            D1 = 1.759; D2 = 6.031; D3 = 0.452; D4 = 1.548;
                            break;
                        case 25:
                            A = 0.600; A1 = 0.919; A2 = 0.153;
                            C2 = 0.9696; C4 = 0.9896;
                            B1 = 0.548; B2 = 1.392; B3 = 0.565; B4 = 1.435;
                            d2 = 3.931; d3 = 0.709;
                            D1 = 1.804; D2 = 6.058; D3 = 0.459; D4 = 1.541;
                            break;
                    }
                    #endregion

                    //Stdev.P 標準差
                    double MeasureAllValueAvg = measureValueList.Average();
                    double sumStdev = 0;
                    for (int i = 0; i < dataCount; i++)
                    {
                        sumStdev += (measureValueList[i] - MeasureAllValueAvg) * (measureValueList[i] - MeasureAllValueAvg);
                    }
                    double StdevS = Math.Sqrt(sumStdev / (dataCount - 1));
                    double StdevN = (DesignUsl - DesignLsl) / 12;

                    if (StdevS == 0) throw new SystemException("數據量不足，請重新選擇!"); ;

                    //計算每個日期的X Bar
                    List<string> tableCalculate = new List<string>(); //尾端的計算欄位，X Bar           
                    List<Quartiles> Quartiles = new List<Quartiles>();
                    tableCalculate.Add("X Bar");
                    foreach (var item1 in sortChartDatas)
                    {
                        double sumX = 0;
                        double XBar = 0;
                        int listCount = 0;
                        List<double> QuartileValue = new List<double>();

                        foreach (var item2 in item1.measureValueList)
                        {
                            QuartileValue.Add(Convert.ToDouble(item2));
                            sumX += Math.Round(Convert.ToDouble(item2), 4);
                            listCount += 1;
                        }
                        XBar = Math.Round(sumX / listCount, 4);
                        tableCalculate.Add(XBar.ToString());
                        Quartiles.Add(CalculateQuartiles(QuartileValue));
                    }
                    tableDatas.Add(tableCalculate);

                    //計算每個日期的R值
                    List<string> tableQuartiles = new List<string>();
                    tableQuartiles.Add("R");
                    double sumR = 0;
                    int countR = 0;
                    foreach (var item in Quartiles)
                    {
                        double max = item.QUpper;
                        double min = item.QLower;
                        double R = max - min;
                        sumR += R;
                        countR += 1;
                        tableQuartiles.Add(Math.Round(R, 5).ToString());
                    }
                    tableDatas.Add(tableQuartiles);

                    //管制界限：上限/下限/中心 R Chart
                    double RCenter = Math.Round(sumR / countR, 5);
                    double RUsl = Math.Round(D4 * RCenter, 5);
                    double RLsl = Math.Round(D3 * RCenter, 5);
                    //管制界限：上限/下限/中心 X Bar
                    double DynamicCenter = Math.Round(MeasureAllValueAvg, 4);
                    double DynamicUsl = Math.Round(DynamicCenter + (A2 * RCenter), 4);
                    double DynamicLsl = Math.Round(DynamicCenter - (A2 * RCenter), 4);

                    //計算製程能力指數：Ca、Cp、Cpk、Pp、Ppk
                    double Ca = Math.Round((DesignValue - DynamicCenter) / ((DesignUsl - DesignLsl) / 2), 4); //Ca = (M-X)/(T/2) M：產品中心位置（規格中心）X：群體的中心（平均值）T：規格寬度（規格上限-規格下限）
                    double Cp = 0; //Cp = T/(6σp)，T= 規格上限(USL) – 規格下限(LSL)、Cp=(USL-X)/(3σp)、Cp=(X-LSL)/(3σp) X為群體中心(平均值)
                    if (DesignUsl == 0) //CpL
                    {
                        Cp = Math.Round((MeasureAllValueAvg - DesignLsl) / (RCenter / d2 * 3), 4);
                    }
                    else if (DesignLsl == 0) //CpU
                    {
                        Cp = Math.Round((DesignUsl - MeasureAllValueAvg) / (RCenter / d2 * 3), 4);
                    }
                    else //CP
                    {
                        Cp = Math.Round((DesignUsl - DesignLsl) / (RCenter / d2 * 6), 4);
                    }
                    double Cpk = Math.Round(Cp * (1 - Math.Abs(Ca)), 4);
                    double Pp = Math.Round((DesignUsl - DesignLsl) / (StdevS * 6), 4);
                    double Ppk = Math.Round(Math.Min((DesignUsl - DynamicCenter) / (StdevS * 3), (DynamicCenter - DesignLsl) / (StdevS * 3)) * Pp, 4);

                    Statistics statistic = new Statistics
                    {
                        dataCount = dataCount.ToString(),
                        DynamicCenter = DynamicCenter.ToString(),
                        DynamicUsl = DynamicUsl.ToString(),
                        DynamicLsl = DynamicLsl.ToString(),
                        n = n.ToString(),
                        GroupCount = GroupCount.ToString(),
                        StdevS = StdevS.ToString(),
                        RCenter = RCenter.ToString(),
                        RUsl = RUsl.ToString(),
                        RLsl = RLsl.ToString(),
                        Ca = Ca.ToString(),
                        Cp = Cp.ToString(),
                        Cpk = Cpk.ToString(),
                        Pp = Pp.ToString(),
                        Ppk = Ppk.ToString()
                    };
                    statistics.Add(statistic);
                    #endregion

                    #region //超規判定及管制圖8種判定規則
                    #region //宣告
                    int One = judgeK[0];
                    int Two = judgeK[1];
                    int Three = judgeK[2];
                    int Four = judgeK[3];
                    int Five = judgeK[4];
                    int Six = judgeK[5];
                    int Seven = judgeK[6];
                    int Eight = judgeK[7];

                    List<JudgeStandards> judgeStandardXBars = new List<JudgeStandards>();
                    List<JudgeStandards> judgeStandardRs = new List<JudgeStandards>();
                    JudgeStandards judgeStandardXBar = new JudgeStandards();
                    JudgeStandards judgeStandardR = new JudgeStandards();
                    bool JudgeToF = false;
                    ArrayList tablefontColors = new ArrayList();
                    List<string> OnePointXBar = new List<string>();
                    List<string> OneValueXBar = new List<string>();
                    List<string> OneDateXBar = new List<string>();
                    List<string> TwoPointXBar = new List<string>();
                    List<string> TwoValueXBar = new List<string>();
                    List<string> TwoDateXBar = new List<string>();
                    List<string> ThreePointXBar = new List<string>();
                    List<string> ThreeValueXBar = new List<string>();
                    List<string> ThreeDateXBar = new List<string>();
                    List<string> FourPointXBar = new List<string>();
                    List<string> FourValueXBar = new List<string>();
                    List<string> FourDateXBar = new List<string>();
                    List<string> FivePointXBar = new List<string>();
                    List<string> FiveValueXBar = new List<string>();
                    List<string> FiveDateXBar = new List<string>();
                    List<string> SixPointXBar = new List<string>();
                    List<string> SixValueXBar = new List<string>();
                    List<string> SixDateXBar = new List<string>();
                    List<string> SevenPointXBar = new List<string>();
                    List<string> SevenValueXBar = new List<string>();
                    List<string> SevenDateXBar = new List<string>();
                    List<string> EightPointXBar = new List<string>();
                    List<string> EightValueXBar = new List<string>();
                    List<string> EightDateXBar = new List<string>();
                    List<string> OnePointR = new List<string>();
                    List<string> OneValueR = new List<string>();
                    List<string> OneDateR = new List<string>();
                    List<string> TwoPointR = new List<string>();
                    List<string> TwoValueR = new List<string>();
                    List<string> TwoDateR = new List<string>();
                    List<string> ThreePointR = new List<string>();
                    List<string> ThreeValueR = new List<string>();
                    List<string> ThreeDateR = new List<string>();
                    List<string> FourPointR = new List<string>();
                    List<string> FourValueR = new List<string>();
                    List<string> FourDateR = new List<string>();
                    List<string> FivePointR = new List<string>();
                    List<string> FiveValueR = new List<string>();
                    List<string> FiveDateR = new List<string>();
                    List<string> SixPointR = new List<string>();
                    List<string> SixValueR = new List<string>();
                    List<string> SixDateR = new List<string>();
                    List<string> SevenPointR = new List<string>();
                    List<string> SevenValueR = new List<string>();
                    List<string> SevenDateR = new List<string>();
                    List<string> EightPointR = new List<string>();
                    List<string> EightValueR = new List<string>();
                    List<string> EightDateR = new List<string>();
                    List<string> dateList = new List<string>();
                    int judgeCount = 0;
                    #endregion

                    foreach (List<string> item in tableDatas)
                    {
                        if (judgeCount == 0)
                        {
                            dateList = item;
                            judgeCount += 1;
                            continue;
                        }
                        List<string> tablefontColor = new List<string>();
                        string colorString = "";
                        int twoPlus = 0;
                        int twoMinus = 0;
                        int threePlus = 0;
                        int threeMinus = 0;
                        int fourPlus = 0;
                        int fourMinus = 0;
                        int fivePlus = 0;
                        int fiveMinus = 0;
                        int sixPlus = 0;
                        int sixMinus = 0;
                        int sevenPlus = 0;
                        int eightPlus = 0;
                        double LastThree = 0;
                        double LastFour = 0;
                        //量測數據
                        for (int i = 1; i < item.Count; i++)
                        {
                            if (item[i] != "")
                            {
                                double tableValue = Convert.ToDouble(item[i]);
                                if (judgeCount < tableDatas.Count - 2)
                                {
                                    if (tableValue < DesignLsl || tableValue > DesignUsl)
                                    {
                                        colorString = "#CE0000";
                                    }
                                    else
                                    {
                                        colorString = "#000000";
                                    }
                                }
                                else if (judgeCount == tableDatas.Count - 2) //X Bar
                                {
                                    if (tableValue < DynamicLsl || tableValue > DynamicUsl)
                                    {
                                        colorString = "##F75000";
                                    }
                                    else
                                    {
                                        colorString = "#000000";
                                    }
                                }
                                else if (judgeCount == tableDatas.Count - 1) //R
                                {
                                    if (tableValue < RLsl || tableValue > RUsl)
                                    {
                                        colorString = "##984B4B";
                                    }
                                    else
                                    {
                                        colorString = "#000000";
                                    }
                                }
                                tablefontColor.Add(colorString);
                            }
                            else
                            {
                                tablefontColor.Add("#000000");
                            }
                        }
                        tablefontColors.Add(tablefontColor);

                        #region //X Bar、R，8種判定規則
                        if (judgeCount >= tableDatas.Count - 2)
                        {
                            //A區為中心+-2σ到中心+-3σ之間
                            //B區為中心+-1σ到中心+-2σ之間
                            //C區為中心到中心+-1σ之間
                            //標準1：1點落在A區外
                            //標準2：連續九點落在同一側
                            //標準3：連續六點遞增或遞減
                            //標準4：14交替，連續十四點，相鄰點上下交錯
                            //標準5：2/3A，連續三點，有兩點落在同一側B區外(A區)
                            //標準6：4/5C，連續五點，有四點落在C區外
                            //標準7：15全C，連續十五點全在兩側C區內
                            //標準8：8缺C，連續八點落在兩側，但沒有任一點在C區
                            //A區: > DynamicCenter +- StdevS * 2;
                            //B區: > DynamicCenter +- StdevS * 1，< DynamicCenter +- StdevS * 2;
                            //C區: > DynamicCenter，< DynamicCenter +- StdevS * 1;
                            for (int i = 1; i < item.Count; i++)
                            {
                                if (item[i] != "")
                                {
                                    double tableValue = Convert.ToDouble(item[i]);
                                    string date = dateList[i].ToString();

                                    if (item[0] == "X Bar")
                                    {
                                        if (One != 0)
                                        {
                                            double StdevRate = Math.Round(StdevS * One, 5);
                                            //邏輯：數值落在A區，就放進LIST
                                            if (tableValue > (DynamicCenter + StdevRate) || tableValue < (DynamicCenter - StdevRate))
                                            {
                                                OnePointXBar.Add(i.ToString());
                                                OneDateXBar.Add(date);
                                                OneValueXBar.Add(tableValue.ToString());
                                            }
                                        }
                                        if (Two != 0)
                                        {
                                            //邏輯1：正與負各自計算，若數據跳到另一邊，則歸0重計
                                            //邏輯2：若plus or minus 小於 Two 則繼續反覆跳
                                            //邏輯3：若滿足 plus or minus 大於等於 Two 的條件，則另一邊不加入TwoJudgeXBar LIST
                                            if (tableValue > DynamicCenter)
                                            {
                                                if (twoMinus < Two)
                                                {
                                                    if (twoPlus == 0)
                                                    {
                                                        TwoPointXBar.Clear();
                                                        TwoValueXBar.Clear();
                                                        TwoDateXBar.Clear();
                                                    }
                                                    TwoPointXBar.Add(i.ToString());
                                                    TwoValueXBar.Add(tableValue.ToString());
                                                    TwoDateXBar.Add(date);
                                                    twoPlus += 1;
                                                    twoMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (twoPlus < Two)
                                                {
                                                    if (twoMinus == 0)
                                                    {
                                                        TwoPointXBar.Clear();
                                                        TwoValueXBar.Clear();
                                                        TwoDateXBar.Clear();
                                                    }
                                                    TwoPointXBar.Add(i.ToString());
                                                    TwoValueXBar.Add(tableValue.ToString());
                                                    TwoDateXBar.Add(date);
                                                    twoMinus += 1;
                                                    twoPlus = 0;
                                                }
                                            }
                                        }
                                        if (Three != 0)
                                        {
                                            //邏輯1：同Two的邏輯，但要與上一個數據對比
                                            if (tableValue >= LastThree)
                                            {
                                                if (threeMinus < Three)
                                                {
                                                    if (threePlus == 0)
                                                    {
                                                        ThreePointXBar.Clear();
                                                        ThreeValueXBar.Clear();
                                                        ThreeDateXBar.Clear();
                                                    }
                                                    ThreePointXBar.Add(i.ToString());
                                                    ThreeValueXBar.Add(tableValue.ToString());
                                                    ThreeDateXBar.Add(date);
                                                    threePlus += 1;
                                                    threeMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (threePlus < Three)
                                                {
                                                    if (threeMinus == 0)
                                                    {
                                                        ThreePointXBar.Clear();
                                                        ThreeValueXBar.Clear();
                                                        ThreeDateXBar.Clear();
                                                    }
                                                    ThreePointXBar.Add(i.ToString());
                                                    ThreeValueXBar.Add(tableValue.ToString());
                                                    ThreeDateXBar.Add(date);
                                                    threeMinus += 1;
                                                    threePlus = 0;
                                                }
                                            }

                                            LastThree = tableValue;
                                        }
                                        if (Four != 0)
                                        {
                                            if (tableValue >= LastFour)
                                            {
                                                if (fourPlus > 0)
                                                {
                                                    fourPlus = 0;
                                                    FourPointXBar.Clear();
                                                    FourValueXBar.Clear();
                                                    FourDateXBar.Clear();
                                                }
                                                FourPointXBar.Add(i.ToString());
                                                FourValueXBar.Add(tableValue.ToString());
                                                FourDateXBar.Add(date);
                                                fourPlus += 1;
                                                fourMinus = 0;
                                            }
                                            else
                                            {
                                                if (fourMinus > 0)
                                                {
                                                    fourMinus = 0;
                                                    FourPointXBar.Clear();
                                                    FourValueXBar.Clear();
                                                    FourDateXBar.Clear();
                                                }
                                                FourPointXBar.Add(i.ToString());
                                                FourValueXBar.Add(tableValue.ToString());
                                                FourDateXBar.Add(date);
                                                fourMinus += 1;
                                                fourPlus = 0;
                                            }
                                            LastFour = tableValue;
                                        }
                                        if (Five != 0)
                                        {
                                            //邏輯1：連續 Five +1 中，有 Five 點超出StdevS * 2(A區)
                                            if (tableValue > (DynamicCenter + StdevS * 2))
                                            {
                                                if (fiveMinus < Five + 1)
                                                {
                                                    if (fivePlus == 0)
                                                    {
                                                        FivePointXBar.Clear();
                                                        FiveValueXBar.Clear();
                                                        FiveDateXBar.Clear();
                                                    }

                                                    FivePointXBar.Add(i.ToString());
                                                    FiveValueXBar.Add(tableValue.ToString());
                                                    FiveDateXBar.Add(date);
                                                    fivePlus += 1;
                                                    fiveMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (DynamicCenter - StdevS * 2))
                                            {
                                                if (fivePlus < Five + 1)
                                                {
                                                    if (fiveMinus == 0)
                                                    {
                                                        FivePointXBar.Clear();
                                                        FiveValueXBar.Clear();
                                                        FiveDateXBar.Clear();
                                                    }
                                                    FivePointXBar.Add(i.ToString());
                                                    FiveValueXBar.Add(tableValue.ToString());
                                                    FiveDateXBar.Add(date);
                                                    fiveMinus += 1;
                                                    fivePlus = 0;
                                                }
                                            }
                                        }
                                        if (Six != 0)
                                        {
                                            if (tableValue > (DynamicCenter + StdevS))
                                            {
                                                if (sixMinus < Six + 1)
                                                {
                                                    if (sixPlus == 0)
                                                    {
                                                        SixPointXBar.Clear();
                                                        SixValueXBar.Clear();
                                                        SixDateXBar.Clear();
                                                    }

                                                    SixPointXBar.Add(i.ToString());
                                                    SixValueXBar.Add(tableValue.ToString());
                                                    SixDateXBar.Add(date);
                                                    sixPlus += 1;
                                                    sixMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (DynamicCenter - StdevS))
                                            {
                                                if (sixPlus < Six + 1)
                                                {
                                                    if (sixMinus == 0)
                                                    {
                                                        SixPointXBar.Clear();
                                                        SixValueXBar.Clear();
                                                        SixDateXBar.Clear();
                                                    }
                                                    SixPointXBar.Add(i.ToString());
                                                    SixValueXBar.Add(tableValue.ToString());
                                                    SixDateXBar.Add(date);
                                                    sixMinus += 1;
                                                    sixPlus = 0;
                                                }
                                            }
                                        }
                                        if (Seven != 0)
                                        {
                                            if (tableValue < DynamicCenter + StdevS && tableValue > DynamicCenter - StdevS)
                                            {
                                                SevenPointXBar.Add(i.ToString());
                                                SevenValueXBar.Add(tableValue.ToString());
                                                SevenDateXBar.Add(date);
                                                sevenPlus += 1;
                                            }
                                            else
                                            {
                                                sevenPlus = 0;
                                                SevenPointXBar.Clear();
                                                SevenValueXBar.Clear();
                                                SevenDateXBar.Clear();
                                            }
                                        }
                                        if (Eight != 0)
                                        {
                                            if (tableValue > DynamicCenter + StdevS && tableValue < DynamicCenter - StdevS)
                                            {
                                                EightPointXBar.Add(i.ToString());
                                                EightValueXBar.Add(tableValue.ToString());
                                                EightDateXBar.Add(date);
                                                eightPlus += 1;
                                            }
                                            else
                                            {
                                                eightPlus = 0;
                                                EightPointXBar.Clear();
                                                EightValueXBar.Clear();
                                                EightDateXBar.Clear();
                                            }
                                        }
                                    }
                                    else if (item[0] == "R")
                                    {
                                        if (One != 0)
                                        {
                                            double StdevRate = Math.Round(StdevS * One, 5);
                                            //邏輯：數值落在A區，就放進LIST
                                            if (tableValue > (RCenter + StdevRate) || tableValue < (RCenter - StdevRate))
                                            {
                                                OnePointR.Add(i.ToString());
                                                OneValueR.Add(tableValue.ToString());
                                                OneDateR.Add(date);
                                            }

                                        }
                                        if (Two != 0)
                                        {
                                            if (tableValue > RCenter)
                                            {
                                                if (twoMinus < Two)
                                                {
                                                    if (twoPlus == 0)
                                                    {
                                                        TwoPointR.Clear();
                                                        TwoValueR.Clear();
                                                        TwoDateR.Clear();
                                                    }
                                                    TwoPointR.Add(i.ToString());
                                                    TwoValueR.Add(tableValue.ToString());
                                                    TwoDateR.Add(date);
                                                    twoPlus += 1;
                                                    twoMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (twoPlus < Two)
                                                {
                                                    if (twoMinus == 0)
                                                    {
                                                        TwoPointR.Clear();
                                                        TwoValueR.Clear();
                                                        TwoDateR.Clear();
                                                    }
                                                    TwoPointR.Add(i.ToString());
                                                    TwoValueR.Add(tableValue.ToString());
                                                    TwoDateR.Add(date);
                                                    twoMinus += 1;
                                                    twoPlus = 0;
                                                }
                                            }
                                        }
                                        if (Three != 0)
                                        {
                                            if (tableValue >= LastThree)
                                            {
                                                if (threeMinus < Three)
                                                {
                                                    if (threePlus == 0)
                                                    {
                                                        ThreePointR.Clear();
                                                        ThreeValueR.Clear();
                                                        ThreeDateR.Clear();
                                                    }
                                                    ThreePointR.Add(i.ToString());
                                                    ThreeValueR.Add(tableValue.ToString());
                                                    ThreeDateR.Add(date);
                                                    threePlus += 1;
                                                    threeMinus = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (threePlus < Three)
                                                {
                                                    if (threeMinus == 0)
                                                    {
                                                        ThreePointR.Clear();
                                                        ThreeValueR.Clear();
                                                        ThreeDateR.Clear();
                                                    }
                                                    ThreePointR.Add(i.ToString());
                                                    ThreeValueR.Add(tableValue.ToString());
                                                    ThreeDateR.Add(date);
                                                    threeMinus += 1;
                                                    threePlus = 0;
                                                }
                                            }

                                            LastThree = tableValue;
                                        }
                                        if (Four != 0)
                                        {
                                            if (tableValue >= LastFour)
                                            {
                                                if (fourPlus > 0)
                                                {
                                                    fourPlus = 0;
                                                    FourPointR.Clear();
                                                    FourValueR.Clear();
                                                    FourDateR.Clear();
                                                }
                                                FourPointR.Add(i.ToString());
                                                FourValueR.Add(tableValue.ToString());
                                                FourDateR.Add(date);
                                                fourPlus += 1;
                                                fourMinus = 0;
                                            }
                                            else
                                            {
                                                if (fourMinus > 0)
                                                {
                                                    fourMinus = 0;
                                                    FourPointR.Clear();
                                                    FourValueR.Clear();
                                                    FourDateR.Clear();
                                                }
                                                FourPointR.Add(i.ToString());
                                                FourValueR.Add(tableValue.ToString());
                                                FourDateR.Add(date);
                                                fourMinus += 1;
                                                fourPlus = 0;
                                            }
                                            LastFour = tableValue;
                                        }
                                        if (Five != 0)
                                        {
                                            if (tableValue > (RCenter + StdevS * 2))
                                            {
                                                if (fiveMinus < Five + 1)
                                                {
                                                    if (fivePlus == 0)
                                                    {
                                                        FivePointR.Clear();
                                                        FiveValueR.Clear();
                                                        FiveDateR.Clear();
                                                    }

                                                    FivePointR.Add(i.ToString());
                                                    FiveValueR.Add(tableValue.ToString());
                                                    FiveDateR.Add(date);
                                                    fivePlus += 1;
                                                    fiveMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (RCenter - StdevS * 2))
                                            {
                                                if (fivePlus < Five + 1)
                                                {
                                                    if (fiveMinus == 0)
                                                    {
                                                        FivePointR.Clear();
                                                        FiveValueR.Clear();
                                                        FiveDateR.Clear();
                                                    }
                                                    FivePointR.Add(i.ToString());
                                                    FiveValueR.Add(tableValue.ToString());
                                                    FiveDateR.Add(date);
                                                    fiveMinus += 1;
                                                    fivePlus = 0;
                                                }
                                            }
                                        }
                                        if (Six != 0)
                                        {
                                            if (tableValue > (RCenter + StdevS))
                                            {
                                                if (sixMinus < Six + 1)
                                                {
                                                    if (sixPlus == 0)
                                                    {
                                                        SixPointR.Clear();
                                                        SixValueR.Clear();
                                                        SixDateR.Clear();
                                                    }

                                                    SixPointR.Add(i.ToString());
                                                    SixValueR.Add(tableValue.ToString());
                                                    SixDateR.Add(date);
                                                    sixPlus += 1;
                                                    sixMinus = 0;
                                                }
                                            }
                                            else if (tableValue < (RCenter - StdevS))
                                            {
                                                if (sixPlus < Six + 1)
                                                {
                                                    if (sixMinus == 0)
                                                    {
                                                        SixPointR.Clear();
                                                        SixValueR.Clear();
                                                        SixDateR.Clear();
                                                    }
                                                    SixPointR.Add(i.ToString());
                                                    SixValueR.Add(tableValue.ToString());
                                                    SixDateR.Add(date);
                                                    sixMinus += 1;
                                                    sixPlus = 0;
                                                }
                                            }
                                        }
                                        if (Seven != 0)
                                        {
                                            if (tableValue < RCenter + StdevS && tableValue > RCenter - StdevS)
                                            {
                                                SevenPointR.Add(i.ToString());
                                                SevenValueR.Add(tableValue.ToString());
                                                SevenDateR.Add(date);
                                                sevenPlus += 1;
                                            }
                                            else
                                            {
                                                sevenPlus = 0;
                                                SevenPointR.Clear();
                                                SevenValueR.Clear();
                                                SevenDateR.Clear();
                                            }
                                        }
                                        if (Eight != 0)
                                        {
                                            if (tableValue > RCenter + StdevS && tableValue < RCenter - StdevS)
                                            {
                                                EightPointR.Add(i.ToString());
                                                EightValueR.Add(tableValue.ToString());
                                                EightDateR.Add(date);
                                                eightPlus += 1;
                                            }
                                            else
                                            {
                                                eightPlus = 0;
                                                EightPointR.Clear();
                                                EightValueR.Clear();
                                                EightDateR.Clear();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        judgeCount += 1;
                    }

                    #region //將判定結果放到list裡
                    //XBar
                    if (OnePointXBar.Count >= 1)
                    {
                        JudgeToF = true;
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 A",
                        Point = OnePointXBar,
                        PointDate = OneDateXBar,
                        PointValue = OneValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Two != 0)
                    {
                        if (TwoPointXBar.Count >= Two)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 B",
                        Point = TwoPointXBar,
                        PointDate = TwoDateXBar,
                        PointValue = TwoValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Three != 0)
                    {
                        if (ThreePointXBar.Count >= Three)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 C",
                        Point = ThreePointXBar,
                        PointDate = ThreeDateXBar,
                        PointValue = ThreeValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Four != 0)
                    {
                        if (FourPointXBar.Count >= Four)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 D",
                        Point = FourPointXBar,
                        PointDate = FourDateXBar,
                        PointValue = FourValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Five != 0)
                    {
                        if (FivePointXBar.Count >= Five)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 E",
                        Point = FivePointXBar,
                        PointDate = FiveDateXBar,
                        PointValue = FiveValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Six != 0)
                    {
                        if (SixPointXBar.Count >= Six)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 F",
                        Point = SixPointXBar,
                        PointDate = SixDateXBar,
                        PointValue = SixValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Seven != 0)
                    {
                        if (SevenPointXBar.Count >= Seven)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 G",
                        Point = SevenPointXBar,
                        PointDate = SevenDateXBar,
                        PointValue = SevenValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);
                    if (Eight != 0)
                    {
                        if (EightPointXBar.Count >= Eight)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardXBar = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 H",
                        Point = EightPointXBar,
                        PointDate = EightDateXBar,
                        PointValue = EightValueXBar,
                    };
                    judgeStandardXBars.Add(judgeStandardXBar);

                    //R
                    if (OnePointR.Count >= 1)
                    {
                        JudgeToF = true;
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 A",
                        Point = OnePointR,
                        PointDate = OneDateR,
                        PointValue = OneValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Two != 0)
                    {
                        if (TwoPointR.Count >= Two)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 B",
                        Point = TwoPointR,
                        PointDate = TwoDateR,
                        PointValue = TwoValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Three != 0)
                    {
                        if (ThreePointR.Count >= Three)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 C",
                        Point = ThreePointR,
                        PointDate = ThreeDateR,
                        PointValue = ThreeValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Four != 0)
                    {
                        if (FourPointR.Count >= Four)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 D",
                        Point = FourPointR,
                        PointDate = FourDateR,
                        PointValue = FourValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Five != 0)
                    {
                        if (FivePointR.Count >= Five)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 E",
                        Point = FivePointR,
                        PointDate = FiveDateR,
                        PointValue = FiveValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Six != 0)
                    {
                        if (SixPointR.Count >= Six)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 F",
                        Point = SixPointR,
                        PointDate = SixDateR,
                        PointValue = SixValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Seven != 0)
                    {
                        if (SevenPointR.Count >= Seven)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 G",
                        Point = SevenPointR,
                        PointDate = SevenDateR,
                        PointValue = SevenValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    if (Eight != 0)
                    {
                        if (EightPointR.Count >= Eight)
                        {
                            JudgeToF = true;
                        }
                        else
                        {
                            JudgeToF = false;
                        }
                    }
                    else
                    {
                        JudgeToF = false;
                    }
                    judgeStandardR = new JudgeStandards
                    {
                        JudgeToF = JudgeToF,
                        whitchJudge = "判定 H",
                        Point = EightPointR,
                        PointDate = EightDateR,
                        PointValue = EightValueR,
                    };
                    judgeStandardRs.Add(judgeStandardR);
                    #endregion

                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result,
                        Quartiles,
                        statistics,
                        tableDatas,
                        tablefontColors,
                        judgeStandardXBars,
                        judgeStandardRs
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

        #region //GetQcType -- 取得送測類型 -- WuTc 2024-09-05
        public string GetQcType()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT QcTypeNo, QcTypeName FROM QMS.QcType";

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

        #endregion

        #region //Update

        #endregion

        #region //Delete

        #endregion

        #region //Calculation
        #region //CalculateQuartiles 計算數字集的四分位數  --  WuTc  --  2024-06-28
        public Quartiles CalculateQuartiles(List<double> numbers)
        {
            var sortedNumbers = numbers.OrderBy(n => n).ToArray();
            int count = sortedNumbers.Length;
            int position25 = count / 4;
            int position50 = count / 2;
            int position75 = 3 * count / 4;

            double Q1 = count % 4 == 1 ? sortedNumbers[position25 - 1] : sortedNumbers[position25];
            double Q2 = count % 4 == 1 ? sortedNumbers[position50 - 1] : sortedNumbers[position50];
            double Q3 = count % 4 == 1 ? sortedNumbers[position75 - 1] : sortedNumbers[position75];

            double upperBound = sortedNumbers[sortedNumbers.Length - 1];
            double lowerBound = sortedNumbers[0];

            Quartiles returnProcess = new Quartiles {
                Q1 = Q1,
                Q2 = Q2,
                Q3 = Q3,
                QUpper = upperBound,
                QLower = lowerBound
            };
            return (returnProcess);
        }
        #endregion
        #endregion
    }
}
