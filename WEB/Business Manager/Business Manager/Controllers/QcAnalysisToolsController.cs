using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using QMSDA;
using System;
using System.Collections.Generic;
using System.Web.Mvc;
using ClosedXML.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;

namespace Business_Manager.Controllers
{
    public class QcAnalysisToolsController : WebController
    {
        private QcAnalysisToolsDA qcAnalysisToolsDA = new QcAnalysisToolsDA();

        #region //View
        public ActionResult QcMeasureMentTracking()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ViewQcAnalysisSPC()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductEnvTestSchedule()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetQcMeasureDataQcItem -- 取得量測紀錄中的量測項目 -- WuTc -- 2024-06-24
        [HttpPost]
        public void GetQcMeasureDataQcItem(int QcItemId = -1, string StartDate = "", string EndDate = "", string MtlItemNo = "", string MtlItemName = "", string QcType = "", string CustomerMtlItemNo = "", string CustomerMtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcMeasureMentTracking", "read");

                if (QcType == "") throw new SystemException("請選擇【量測類別】!");
                if (MtlItemName == "" && MtlItemNo == "") throw new SystemException("請輸入【品號】或【品名】!");

                #region //Request
                qcAnalysisToolsDA = new QcAnalysisToolsDA();
                dataRequest = qcAnalysisToolsDA.GetQcMeasureDataQcItem(QcItemId, StartDate, EndDate, MtlItemNo, MtlItemName, QcType, CustomerMtlItemNo, CustomerMtlItemName
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMeasureDataforSPC -- 取得SPC所需的量測資料 -- WuTc -- 2024-06-24
        public class JudgeStandard
        {
            public int One { get; set; }
            public int Two { get; set; }
            public int Three { get; set; }
            public int Four { get; set; }
            public int Five { get; set; }
            public int Six { get; set; }
            public int Seven { get; set; }
            public int Eight { get; set; }
        }
        [HttpPost]
        public void GetQcMeasureDataforSPC(JudgeStandard judgeK, int QcItemId = -1, string StartDate = "", string EndDate = "", string MtlItemNo = "", string MtlItemName = "", string QcType = "", string CustomerMtlItemNo = "", string CustomerMtlItemName = "", double DesignValue = 0
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcMeasureMentTracking", "read");
                //8項判定標準
                List<int> judge = new List<int>();
                judge.Add(judgeK.One);
                judge.Add(judgeK.Two);
                judge.Add(judgeK.Three);
                judge.Add(judgeK.Four);
                judge.Add(judgeK.Five);
                judge.Add(judgeK.Six);
                judge.Add(judgeK.Seven);
                judge.Add(judgeK.Eight);
                #region //Request
                qcAnalysisToolsDA = new QcAnalysisToolsDA();
                dataRequest = qcAnalysisToolsDA.GetQcMeasureDataforSPC(judge, QcItemId, StartDate, EndDate, MtlItemNo, MtlItemName, QcType, CustomerMtlItemNo, CustomerMtlItemName, DesignValue
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcType 取得量測類型資料
        [HttpPost]
        public void GetQcType()
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcAnalysisToolsDA = new QcAnalysisToolsDA();
                dataRequest = qcAnalysisToolsDA.GetQcType();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion

        #region //Add

        #endregion

        #region //Update

        #endregion

        #region //Delete

        #endregion

        #region //Calculationthrow new
        #region //SPCforJudgeStandard -- 8項判定標準 -- WuTc -- 2024-07-08

        [HttpPost]
        public void SPCforJudgeStandard(JudgeStandard judgeK, int QcItemId = -1, string StartDate = "", string EndDate = "", string MtlItemNo = "", string MtlItemName = "", string QcType = "", string CustomerMtlItemNo = "", string CustomerMtlItemName = ""
            , double DesignValue = 0, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcMeasureMentTracking", "read");
                List<int> judge = new List<int>();
                judge.Add(judgeK.One);
                judge.Add(judgeK.Two);
                judge.Add(judgeK.Three);
                judge.Add(judgeK.Four);
                judge.Add(judgeK.Five);
                judge.Add(judgeK.Six);
                judge.Add(judgeK.Seven);
                judge.Add(judgeK.Eight);
                #region //Request
                qcAnalysisToolsDA = new QcAnalysisToolsDA();
                dataRequest = qcAnalysisToolsDA.GetQcMeasureDataforSPC(judge, QcItemId, StartDate, EndDate, MtlItemNo, MtlItemName, QcType, CustomerMtlItemNo, CustomerMtlItemName, DesignValue
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion

        #region //Excel
        #region //SPCOutputForExcel -- SPC資料匯出至excel -- WuTc -- 2024-07-09
        [HttpPost]
        public void SPCOutputForExcel(JudgeStandard judgeK, int QcItemId = -1, string StartDate = "", string EndDate = "", string MtlItemNo = "", string MtlItemName = "", string QcType = "", string CustomerMtlItemNo = "", string CustomerMtlItemName = "", double DesignValue = 0, string XBarChart = "", string RChart = "", string Pability = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcMeasureMentTracking", "read");
                //8項判定標準
                List<int> judge = new List<int>();
                judge.Add(judgeK.One);
                judge.Add(judgeK.Two);
                judge.Add(judgeK.Three);
                judge.Add(judgeK.Four);
                judge.Add(judgeK.Five);
                judge.Add(judgeK.Six);
                judge.Add(judgeK.Seven);
                judge.Add(judgeK.Eight);
                #region //Request
                qcAnalysisToolsDA = new QcAnalysisToolsDA();
                dataRequest = qcAnalysisToolsDA.GetQcMeasureDataforSPCExcel(judge, QcItemId, StartDate, EndDate, MtlItemNo, MtlItemName, QcType, CustomerMtlItemNo, CustomerMtlItemName, DesignValue
                    , OrderBy, PageIndex, PageSize);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式
                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.None;
                    defaultStyle.Border.TopBorderColor = XLColor.NoColor;
                    defaultStyle.Border.BottomBorderColor = XLColor.NoColor;
                    defaultStyle.Border.LeftBorderColor = XLColor.NoColor;
                    defaultStyle.Border.RightBorderColor = XLColor.NoColor;
                    defaultStyle.Fill.BackgroundColor = XLColor.NoColor;
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    defaultStyle.Font.FontSize = 12;
                    defaultStyle.Font.Bold = false;
                    defaultStyle.Protection.SetLocked(false);
                    #endregion

                    #region //titleStyle
                    var titleStyle = XLWorkbook.DefaultStyle;
                    titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    titleStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    titleStyle.Border.TopBorderColor = XLColor.Black;
                    titleStyle.Border.BottomBorderColor = XLColor.Black;
                    titleStyle.Border.LeftBorderColor = XLColor.Black;
                    titleStyle.Border.RightBorderColor = XLColor.Black;
                    titleStyle.Fill.BackgroundColor = XLColor.NoColor;
                    titleStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    titleStyle.Font.FontSize = 16;
                    titleStyle.Font.Bold = false;
                    #endregion

                    #region //headerStyle
                    var headerStyle = XLWorkbook.DefaultStyle;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    headerStyle.Border.TopBorderColor = XLColor.Black;
                    headerStyle.Border.BottomBorderColor = XLColor.Black;
                    headerStyle.Border.LeftBorderColor = XLColor.Black;
                    headerStyle.Border.RightBorderColor = XLColor.Black;
                    headerStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dateStyle.Border.TopBorder = XLBorderStyleValues.None;
                    dateStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    dateStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    dateStyle.Border.RightBorder = XLBorderStyleValues.None;
                    dateStyle.Border.TopBorderColor = XLColor.NoColor;
                    dateStyle.Border.BottomBorderColor = XLColor.NoColor;
                    dateStyle.Border.LeftBorderColor = XLColor.NoColor;
                    dateStyle.Border.RightBorderColor = XLColor.NoColor;
                    dateStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dateStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    dateStyle.Font.FontSize = 12;
                    dateStyle.Font.Bold = false;
                    dateStyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    #endregion

                    #region //numberStyle
                    var numberStyle = XLWorkbook.DefaultStyle;
                    numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    numberStyle.Border.TopBorder = XLBorderStyleValues.None;
                    numberStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    numberStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    numberStyle.Border.RightBorder = XLBorderStyleValues.None;
                    numberStyle.Border.TopBorderColor = XLColor.NoColor;
                    numberStyle.Border.BottomBorderColor = XLColor.NoColor;
                    numberStyle.Border.LeftBorderColor = XLColor.NoColor;
                    numberStyle.Border.RightBorderColor = XLColor.NoColor;
                    numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                    numberStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "_-* #,##0.00_-;-* #,##0.00_-;_-* \" - \"??_-;_-@_-";
                    #endregion

                    #region //currencyStyle
                    var currencyStyle = XLWorkbook.DefaultStyle;
                    currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    currencyStyle.Border.TopBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.BottomBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.LeftBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.RightBorder = XLBorderStyleValues.None;
                    currencyStyle.Border.TopBorderColor = XLColor.NoColor;
                    currencyStyle.Border.BottomBorderColor = XLColor.NoColor;
                    currencyStyle.Border.LeftBorderColor = XLColor.NoColor;
                    currencyStyle.Border.RightBorderColor = XLColor.NoColor;
                    currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                    currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.DateFormat.Format = "$#,##0.00";
                    #endregion

                    #region //tableHeaderStyle
                    var tableHeaderStyle = XLWorkbook.DefaultStyle;
                    tableHeaderStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    tableHeaderStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    tableHeaderStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    tableHeaderStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    tableHeaderStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    tableHeaderStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    tableHeaderStyle.Border.TopBorderColor = XLColor.Black;
                    tableHeaderStyle.Border.BottomBorderColor = XLColor.Black;
                    tableHeaderStyle.Border.LeftBorderColor = XLColor.Black;
                    tableHeaderStyle.Border.RightBorderColor = XLColor.Black;
                    tableHeaderStyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    tableHeaderStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    tableHeaderStyle.Font.FontSize = 14;
                    tableHeaderStyle.Font.Bold = true;
                    #endregion

                    #region //tableStyle
                    var tableStyle = XLWorkbook.DefaultStyle;
                    tableStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    tableStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    tableStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    tableStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    tableStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    tableStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    tableStyle.Border.TopBorderColor = XLColor.Black;
                    tableStyle.Border.BottomBorderColor = XLColor.Black;
                    tableStyle.Border.LeftBorderColor = XLColor.Black;
                    tableStyle.Border.RightBorderColor = XLColor.Black;
                    tableStyle.Fill.BackgroundColor = XLColor.WhiteSmoke;
                    tableStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    tableStyle.Font.FontSize = 12;
                    tableStyle.Font.Bold = true;
                    #endregion
                    #endregion

                    #region //參數初始化
                    dynamic[] result = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] statistics = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["statistics"].ToString());
                    dynamic[] Quartiles = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["Quartiles"].ToString());
                    dynamic[] judgeStandardXBars = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["judgeStandardXBars"].ToString());
                    dynamic[] judgeStandardRs = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["judgeStandardRs"].ToString());
                    string tableDatas = JObject.Parse(dataRequest)["tableDatas"].ToString();
                    string[] spDatas = Regex.Split(tableDatas.Replace("\r\n", "").Replace("\"", "").Replace("[", "").Replace(" ", ""), "]");

                    string[] header = new string[] { "品號", "品名", "項目編號", "球標", "項目別稱", "設計值", "上公差", "下公差", "單位", "量測儀器", "時間區間" };
                    string[] header2 = new string[] { "組數", "N", "s", "管制中心(XBar)", "管制上限(XBar)", "管制下限(XBar)", "管制中心(R)", "管制上限(R)", "管制下限(R)", "Ca", "Cp", "Cpk", "Pp", "Ppk" };
                    string headerCellIndex = "";
                    int rowIndex = 1;
                    int colIndex = 1;

                    byte[] excelFile;
                    string excelFileName = "SPC_" + MtlItemName + "_" + result[0].QcItemDesc + "_" + StartDate + "-" + EndDate;
                    string excelsheetName = "SPC";

                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //表頭圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var logoimage = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 15).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            headerCellIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(headerCellIndex).Value = header[i];
                            worksheet.Cell(headerCellIndex).Style = headerStyle;
                        }
                        //value
                        rowIndex += 1;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Style = tableStyle;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = result[0].MtlItemNo.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = result[0].MtlItemName.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = result[0].QcItemNo.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = result[0].BallMark.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = result[0].QcItemDesc.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = result[0].DesignValue.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = result[0].UpperTolerance.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = result[0].LowerTolerance.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = result[0].Unit.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = result[0].MachineName.ToString() + "(" + result[0].MachineDesc.ToString() + ")";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = StartDate.ToString() + " - " + EndDate.ToString();

                        rowIndex += 1;
                        for (int i = 0; i < header2.Length; i++)
                        {
                            headerCellIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(headerCellIndex).Value = header2[i];
                            worksheet.Cell(headerCellIndex).Style = headerStyle;
                        }
                        //value
                        rowIndex += 1;
                        worksheet.Range(rowIndex, 1, rowIndex, 14).Style = tableStyle;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = statistics[0].GroupCount.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = statistics[0].n.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = statistics[0].StdevS.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = statistics[0].DynamicCenter.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = statistics[0].DynamicUsl.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = statistics[0].DynamicLsl.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = statistics[0].RCenter.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = statistics[0].RUsl.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = statistics[0].RLsl.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = statistics[0].Ca.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = statistics[0].Cp.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = statistics[0].Cpk.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = statistics[0].Pp.ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = statistics[0].Ppk.ToString();
                        #endregion

                        rowIndex += 2;
                        #region //BODY
                        #region //資料表格
                        foreach (var item in spDatas)
                        {
                            colIndex = 1;
                            if (item != "")
                            {
                                rowIndex++;
                                string[] ListDatas = item.Split(',');
                                if (ListDatas.ToString() != "")
                                {
                                    for (int i = 0; i < ListDatas.Length; i++)
                                    {                                        
                                        if (item.IndexOf('<') != -1)
                                        {
                                            string[] spListDatas = ListDatas[i].Split('(');
                                            if (ListDatas[i].IndexOf("<br>") != -1)
                                            {
                                                spListDatas = ListDatas[i].Replace("<br>", "").Split('(');
                                            }
                                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex)).Style = tableHeaderStyle;
                                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex)).Value = spListDatas[0].ToString();
                                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex + 1)).Style = tableHeaderStyle;
                                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex + 1)).Value = spListDatas[1].Replace(")", "").ToString();
                                        }
                                        else
                                        {
                                            //跳過首尾空白
                                            if (i == 0)
                                            {
                                                continue;
                                            }

                                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex)).Style = tableStyle;
                                            if (ListDatas[i].ToString() != "")
                                            {
                                                worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex)).SetValue(ListDatas[i].ToString()).DataType = XLDataType.Text;
                                            }
                                        }
                                        colIndex++;
                                    }
                                }
                                if (item.IndexOf('<') != -1)
                                {
                                    rowIndex += 1;
                                }
                            }
                        }

                        rowIndex++;
                        worksheet.Range(rowIndex, 1, rowIndex + 4, 1).Style = tableStyle;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).SetValue("25%");
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex + 1)).SetValue("50%");
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex + 2)).SetValue("75%");
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex + 3)).SetValue("Min");
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex + 4)).SetValue("Max");

                        colIndex = 2;
                        foreach (var item in Quartiles)
                        {
                            worksheet.Range(rowIndex, colIndex, rowIndex + 4, colIndex).Style = tableStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex)).SetValue(item.Q1.ToString());
                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex + 1)).SetValue(item.Q2.ToString());
                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex + 2)).SetValue(item.Q3.ToString());
                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex + 3)).SetValue(item.QLower.ToString());
                            worksheet.Cell(BaseHelper.MergeNumberToChar(colIndex, rowIndex + 4)).SetValue(item.QUpper.ToString());
                            colIndex++;
                        }

                        rowIndex += 7;
                        worksheet.Range(rowIndex - 1, 1, rowIndex + 7, 1).Style = tableHeaderStyle;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex - 1)).SetValue("X Bar 8項判定");
                        foreach (var item in judgeStandardXBars)
                        {
                            worksheet.Range(rowIndex, 2, rowIndex, 5).Style = tableStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).SetValue(item.whitchJudge.ToString());
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).SetValue(item.JudgeToF.ToString());
                            int j = 0;
                            foreach(var item2 in item.PointDate)
                            {
                                if (item2.ToString() != "")
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3 + j, rowIndex)).Style = tableStyle;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3 + j, rowIndex)).SetValue(item2.ToString() + "\r\n" + item.PointValue[j].ToString());
                                }
                                j++;
                            }
                            rowIndex++;
                        }

                        rowIndex += 2;
                        worksheet.Range(rowIndex - 1, 1, rowIndex + 7, 1).Style = tableHeaderStyle;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex - 1)).SetValue("R Chart 8項判定");
                        foreach (var item in judgeStandardRs)
                        {
                            worksheet.Range(rowIndex, 2, rowIndex, 5).Style = tableStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).SetValue(item.whitchJudge.ToString());
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).SetValue(item.JudgeToF.ToString());
                            int j = 0;
                            foreach (var item2 in item.PointDate)
                            {
                                if (item2.ToString() != "")
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3 + j, rowIndex)).Style = tableStyle;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3 + j, rowIndex)).SetValue(item2.ToString() + "\r\n" + item.PointValue[j].ToString());
                                }
                                j++;
                            }
                            rowIndex++;
                        }
                        #endregion
                        #endregion

                        #region //圖片
                        #region //將 Base64Image 解碼為byte[]再轉成image
                        // 解码Base64字符串为字节数组
                        string Base64DataReProcess(string base64string)
                        {
                            string dummyData = base64string.Trim().Replace("%", "").Replace("data:image/png;base64", "").Replace(",", "").Replace(" ", "+");
                            if (dummyData.Length % 4 > 0)
                            {
                                dummyData = dummyData.PadRight(dummyData.Length + 4 - dummyData.Length % 4, '=');
                            }
                            return dummyData;
                        };

                        //put pictures
                        string imageTransition(byte[] imageBytes, string chart, int row)
                        {
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex - 1)).Value = chart;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex - 1)).Style = tableHeaderStyle;
                            using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
                            {
                                ms.Write(imageBytes, 0, imageBytes.Length);
                                worksheet.Row(row).Height = 400;
                                worksheet.Range(row, 1, row, 13).Merge().Style = tableStyle;
                                worksheet.AddPicture(ms).MoveTo(worksheet.Cell(row, 2)).Scale(0.85);
                            }
                            return "OK";
                        };
                        #endregion

                        rowIndex += 2;
                        //X Bar Chart
                        if (XBarChart != "")
                        {
                            byte[] imageXBarBytes = Convert.FromBase64String(Base64DataReProcess(XBarChart));
                            var XbarImage = imageTransition(imageXBarBytes, "X Bar", rowIndex);
                            rowIndex += 3;
                        }
                        //R Chart
                        if (RChart != "")
                        {
                            byte[] imageRBytes = Convert.FromBase64String(Base64DataReProcess(Base64DataReProcess(RChart)));
                            var RImage = imageTransition(imageRBytes, "R Chart", rowIndex);
                            rowIndex += 3;
                        }
                        //Process ability Chart
                        if (Pability != "")
                        {
                            byte[] imagePabilityBytes = Convert.FromBase64String(Base64DataReProcess(Pability));
                            var PabilityImage = imageTransition(imagePabilityBytes, "Process ability", rowIndex);
                            rowIndex += 3;
                        }
                        #endregion
                        #region //設定
                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Columns().Width = 21;
                        #endregion

                        #region //置中
                        //worksheet.Columns("A:C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");
                        //worksheet.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        //worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        //worksheet.Cell(4, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(3, 1, 8, 1).Style.Protection.SetLocked(true);
                        //worksheet.Range(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
                        #endregion

                        #region //合併儲存格
                        //worksheet.Range(8, 1, 8, 9).Merge();
                        #endregion

                        #region //公式
                        ////R=Row、C=Column
                        ////[]指相對位置，[-]向前，[+]向後
                        ////未[]指絕對位置
                        //worksheet.Cell(10, 3).FormulaR1C1 = "RC[-2]+R3C2";
                        //worksheet.Cell(11, 3).FormulaR1C1 = "RC[-2]+RC[2]";
                        //worksheet.Cell(12, 3).FormulaR1C1 = "RC[-2]+RC2";
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(9);
                        #endregion

                        #region //隱藏
                        //worksheet.Column(1).Hide();
                        #endregion

                        #region //格式化
                        //string conditionalFormatRange = "G:G";
                        //string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        //{
                        //    range.AddConditionalFormat().WhenEquals("M")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.BabyBlue);

                        //    range.AddConditionalFormat().WhenEquals("F")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.Red);
                        //}

                        //string conditionalFormatRange2 = "E:E";
                        //foreach (var range in worksheet.Ranges(conditionalFormatRange2))
                        //{
                        //    range.AddConditionalFormat().WhenContains("1")
                        //        //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                        //        .Font.SetFontColor(XLColor.RedNcs);
                        //}
                        #endregion

                        #region //複製
                        //worksheet.CopyTo("複製");
                        #endregion
                        #endregion

                        #region //設定刻號欄寬
                        worksheet.Column(11).Width = 50;
                        #endregion

                        #region //表格
                        //var fakeData = Enumerable.Range(1, 5)
                        //.Select(x => new FakeData
                        //{
                        //    Time = TimeSpan.FromSeconds(x * 123.667),
                        //    X = x,
                        //    Y = -x,
                        //    Address = "a" + x,
                        //    Distance = x * 100
                        //}).ToArray();

                        //var table = worksheet.Cell(10, 1).InsertTable(fakeData);
                        //table.Style.Font.FontSize = 9;
                        //var tableData = worksheet.Cell(17, 1).InsertData(fakeData);
                        //tableData.Style.Font.FontSize = 9;
                        //worksheet.Range(11, 1, 21, 1).Style.DateFormat.Format = "HH:mm:ss.000";
                        #endregion

                        #region //EXCEL匯出
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);

                            excelFile = output.ToArray();
                        }
                        #endregion
                    }

                    string fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = excelFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = excelFileName,
                        fileExtension = ".xlsx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
                }

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion
    }
}