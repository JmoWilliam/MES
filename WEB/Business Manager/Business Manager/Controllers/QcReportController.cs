using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using QMSDA;
using Newtonsoft.Json.Linq;
using System.Reflection;
using Newtonsoft.Json;
using ClosedXML.Excel;
using System.IO;
using Helpers;
using System.Text;
using System.Drawing;
using NiceLabel.SDK;
using System.Text.RegularExpressions;
using System.IO.Ports;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Business_Manager.Controllers
{
    public class QcReportController : WebController
    {
        private QcReportDA qcReportDA = new QcReportDA();

        #region //View
        public ActionResult QcDeliveryReport()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult QcDeliveryReportModify()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult QcGoodsReceiptReport()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetDeliveryDetail -- 取得出貨單身紀錄 -- WuTc -- 2024-05-27
        [HttpPost]
        public void GetDeliveryDetail(string StartDate = "", string EndDate = "", int CustomerId = -1, int DcId = -1, string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDeliveryReport", "read");

                #region //Request
                qcReportDA = new QcReportDA();
                dataRequest = qcReportDA.GetDelveryDetail(StartDate, EndDate, CustomerId, DcId, MtlItemNo, MtlItemName
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

        #region //GetInventoryDetail -- 取得入庫單紀錄 -- WuTc -- 2024-07-23
        [HttpPost]
        public void GetInventoryDetail(string StartDate = "", string EndDate = "", int CustomerId = -1, string MoErpNo = "", string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDeliveryReport", "read");

                #region //Request
                qcReportDA = new QcReportDA();
                dataRequest = qcReportDA.GetInventoryDetail(StartDate, EndDate, CustomerId, MoErpNo, MtlItemNo, MtlItemName
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

        #region //GetModifyMeasureData -- 取得修改後上傳的excel筆數 -- WuTc -- 2024-08-09
        [HttpPost]
        public void GetModifyMeasureData(string StartDate = "", string EndDate = "", string MoIdErpNo = "", string MtlItemNo = "", string MtlItemName = "", string UserNo = "", int CustomerId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDeliveryReportModify", "read");

                #region //Request
                qcReportDA = new QcReportDA();
                dataRequest = qcReportDA.GetModifyMeasureData(StartDate, EndDate, MoIdErpNo, MtlItemNo, MtlItemName, UserNo, CustomerId
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

        #region //GetGoodsReceiptReportDataList -- 取得進貨檢驗量測紀錄列表 -- WuTc -- 2024-10-05
        [HttpPost]
        public void GetGoodsReceiptReportDataList(string StartDate = "", string EndDate = "", int SupplierId = -1, string GrFullErpNo = "", string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDeliveryReport", "read");

                #region //Request
                qcReportDA = new QcReportDA();
                dataRequest = qcReportDA.GetGoodsReceiptReportDataList(StartDate, EndDate, SupplierId, GrFullErpNo, MtlItemNo, MtlItemName
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

        #region //GetSupplier -- 取得供應商資料 -- WuTc -- 2024-10-05
        [HttpPost]
        public void GetSupplier(string Status)
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptReport", "read");

                #region //Request
                qcReportDA = new QcReportDA();
                dataRequest = qcReportDA.GetSupplier(Status);
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

        #region //DownLoad
        #region //Excel
        #region //DeliveryReportOutputForExcel 出貨報告產出excel 主程式 -- WuTc -- 2024-05-22
        public void DeliveryReportOutputForExcel(string DataId = "", string SearchDate = "", string MtlItemNo = "", string MtlItemName = "", int CustomerId = -1, int ProductType = -1, string reportType = "")
        {
            try
            {
                WebApiLoginCheck("QcDeliveryReport", "excel");

                JObject dataRequestJson = new JObject();
                object result = "";
                int count = 0;
                #region //確認是否有相對應報表格式

                if (CustomerId > 0)
                {
                    dataRequest = qcReportDA.GetCustomerReport(CustomerId, ProductType);
                    dataRequestJson = JObject.Parse(dataRequest);
                    count = dataRequestJson["data"].Count();
                }
                string ExcelempryFileName = "";
                string encoding = "UTF-8";
                Type thisType = this.GetType();

                //若 SCM.DeliveryCustomerReport 找不到  FunctionName ，則呼叫制式 DeliveryReportOutputForExcelStandard
                if (count > 0)
                {
                    foreach (var item in dataRequestJson["data"])
                    {
                        MethodInfo methodInfo = thisType.GetMethod(item["FunctionName"].ToString());
                        ExcelempryFileName = item["EmptyReportName"].ToString();
                        encoding = item["Encoding"].ToString();
                        object[] parameters = new object[] { DataId, SearchDate, MtlItemNo, MtlItemName, ExcelempryFileName, encoding, reportType };
                        methodInfo.Invoke(this, parameters);
                    }
                }
                else
                {
                    MethodInfo methodInfo = thisType.GetMethod("DeliveryReportOutputForExcelStandard");
                    ExcelempryFileName = "Quality_Inspection_Standard_Shipping_Report.xlsx";
                    //先預設為舜宇的報告
                    //MethodInfo methodInfo = thisType.GetMethod("DeliveryReportOutputForExcelSU");
                    //ExcelempryFileName = "Quality_Inspection_CustomerSU_Shipping_Report.xlsx";
                    object[] parameters = new object[] { DataId, SearchDate, MtlItemNo, MtlItemName, ExcelempryFileName, encoding, reportType };
                    methodInfo.Invoke(this, parameters);
                }
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

        #region //DeliveryReportOutputForExcelStandard 出貨報告產出excel 制式表格 -- WuTc -- 2024-05-28
        public void DeliveryReportOutputForExcelStandard(string DataId = "", string SearchDate = "", string MtlItemNo = "", string MtlItemName = "", string emptyFileName = "", string encoding = "", string reportType = "")
        {
            try
            {
                emptyFileName = "Quality_Inspection_Standard_Lens_Report.xlsx";
                #region //Request
                qcReportDA = new QcReportDA();
                if (reportType == "M")
                {
                    dataRequest = qcReportDA.GetDeliveryMeasurementDataModify(DataId, SearchDate, MtlItemNo, MtlItemName, "", -1, -1); //修改後的報告資料
                }
                else
                {
                    dataRequest = qcReportDA.GetDeliveryMeasurementData(DataId, SearchDate, MtlItemNo, MtlItemName, "", -1, -1);
                }
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    if (data.Count() > 0)
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
                        dateStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.TopBorderColor = XLColor.Black;
                        dateStyle.Border.BottomBorderColor = XLColor.Black;
                        dateStyle.Border.LeftBorderColor = XLColor.Black;
                        dateStyle.Border.RightBorderColor = XLColor.Black;
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
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
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
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.DateFormat.Format = "$#,##0.00";
                        #endregion

                        #region //customizedStyle
                        List<XLColor> xLColors = new List<XLColor>
                    {
                        XLColor.TeaRoseRose,
                        XLColor.PeachOrange,
                        XLColor.Flavescent,
                        XLColor.MediumSpringBud,
                        XLColor.EtonBlue,
                        XLColor.MossGreen,
                        XLColor.LightBlue,
                        XLColor.WildBlueYonder,
                        XLColor.Wisteria,
                        XLColor.PinkPearl
                    };

                        List<XLBorderStyleValues> xLBorderSizes = new List<XLBorderStyleValues>
                    {
                        XLBorderStyleValues.Thin,
                        XLBorderStyleValues.Thick
                    };

                        List<XLColor> xLBorderColors = new List<XLColor>
                    {
                        XLColor.Black,
                        XLColor.CornflowerBlue
                    };
                        #endregion
                        #endregion

                        #region //參數初始化

                        byte[] excelFile;
                        string excelFileName = "";
                        string DefaultStName = "MeasureReport";
                        string DefaultpicStName = "SurfaceAccuracy";

                        int colIndex = 12;
                        int rowIndex = 12;
                        #endregion

                        #region //EXCEL
                        using (var workbook = new XLWorkbook(Server.MapPath("~/PdfTemplate/MES/QcReport/Shipment/" + emptyFileName)))
                        {
                            IXLWorksheet firstsheet = workbook.Worksheet(DefaultStName + "-1.1");
                            IXLWorksheet imageSheet = workbook.Worksheet(DefaultpicStName + "-1");

                            #region //圖片
                            //var imagePath = Server.MapPath("~/Content/images/logo.png");
                            //var image = firstsheet.AddPicture(imagePath).MoveTo(firstsheet.Cell(1, 2)).Scale(0.07);
                            //firstsheet.Row(rowIndex).Height = 40;
                            #endregion

                            #region //HEADER

                            #endregion

                            #region //BODY
                            string lastItemNo = "";
                            int serialCode = 0;
                            int rowSheetCount = 1;
                            int colSheetCount = 2;
                            string CurrentSheetName = DefaultStName + "-1.1";
                            string measureMachine = "";
                            int itemCount = 0;
                            int letteringCount = 0;
                            int totalLettering = 0;
                            int calcColSheet = -1;
                            string QcRecordId = "";
                            string inputType = "";
                            string mtlItemId = "";
                            string MoId = "";
                            List<string> pcsNumberList = new List<string> { };

                            foreach (var item in data)
                            {
                                #region //宣告                                
                                string qcItemNo = item.QcItemNo.ToString();
                                string itemNo = item.QcItemNo + item.MachineNumber;
                                string MachineNo = item.MachineNo;
                                string MachineName = item.MachineName;
                                string ZAxis = item.ZAxis == 0 ? null : item.ZAxis;
                                string Surveyor = item.QcUserName;
                                string CusShortName = item.CustomerShortName;
                                string MweErpNo = item.MweErpNo;
                                string QcClassName = item.QcClassName;
                                string CustomerMtlItemNo = item.CustomerMtlItemNo;
                                string CustomerDwgNo = item.CustomerDwgNo;
                                string mtlItemName = item.MtlItemName;
                                string abnormalNo = item.AbnormalqualityNo;
                                string QcItemId = item.QcItemId.ToString();
                                string QmmDetailId = item.QmmDetailId.ToString();
                                string BarcodeId = item.QcBarcodeId.ToString();
                                string makeCount = item.MakeCount.ToString();
                                MoId = item.MoId.ToString();
                                mtlItemId = item.MtlItemId.ToString();
                                string pcsNumber = "";
                                int totalQuantity = item.Quantity;
                                string Confirmer = "";
                                inputType = item.InputType;
                                QcRecordId = item.QcRecordId;

                                string MeasureValue = "";
                                DateTime DocDate;
                                string serialNo = "";
                                string machineNo = "";
                                string lettering = "";

                                //QcItemDesc、DesignValue、UpperTolerance、LowerTolerance、BallMark、Unit 這六項規格，若允收標準有，則用允收標準的
                                string QcItemDesc = "";
                                string DesignValue = "";
                                string UpperTolerance = "";
                                string LowerTolerance = "";
                                string BallMark = "";
                                string Unit = "";

                                string mtlQcItemDesc = item.MtlQcItemDesc;
                                string mtlDesignValue = item.MtlDesignValue;
                                string mtlUpperTolerance = item.MtlUpperTolerance;
                                string mtlLowerTolerance = item.MtlLowerTolerance;
                                string mtlBallMark = item.MtlBallMark;
                                string mtlUnit = item.MtlUnit;

                                //若允收標準有維護項目規格，則報告上帶出允收標準的規格
                                if (mtlQcItemDesc != null) { QcItemDesc = mtlQcItemDesc; } else { QcItemDesc = item.QcItemDesc; }
                                if (mtlDesignValue != null) { DesignValue = mtlDesignValue; } else { QcItemDesc = item.DesignValue; }
                                if (mtlUpperTolerance != null) { UpperTolerance = mtlUpperTolerance; } else { UpperTolerance = item.UpperTolerance; }
                                if (mtlLowerTolerance != null) { LowerTolerance = mtlLowerTolerance; } else { LowerTolerance = item.LowerTolerance; }
                                if (mtlBallMark != null) { BallMark = mtlBallMark; } else { BallMark = item.BallMark; }
                                if (mtlUnit != null) { Unit = mtlUnit; } else { Unit = item.Unit; }
                                //設計值DesignValue 若為數字，上下公差若為空值，則改為0
                                if (decimal.TryParse(DesignValue, out decimal s) != false)
                                {
                                    if (UpperTolerance == "") UpperTolerance = "0";
                                    if (LowerTolerance == "") LowerTolerance = "0";
                                }

                                if (reportType == "M")
                                {
                                    machineNo = MachineNo;
                                    lettering = item.LetteringSeq; //刻號
                                    MeasureValue = item.ModifyValue;
                                    DocDate = item.CreateDate;
                                    Surveyor = item.Surveyor;
                                    Confirmer = item.Confirmer;
                                }
                                else
                                {
                                    if (inputType == "Lettering")
                                    {
                                        pcsNumber = item.ItemSeq; //刻號
                                    }
                                    else if (inputType == "Cavity")
                                    {
                                        pcsNumber = item.Cavity; //穴號
                                    }
                                    else if (inputType == "LotNumber")
                                    {
                                        pcsNumber = item.LotNumber; //批號
                                    }
                                    else if (inputType == "BarcodeId")
                                    {
                                        pcsNumber = item.BarcodeId; //條碼
                                    }

                                    machineNo = QmmDetailId;
                                    lettering = pcsNumber + " (" + BarcodeId + ")";
                                    MeasureValue = item.MeasureValue;
                                    DocDate = item.DocDate;
                                    Surveyor = item.QcUserName;
                                }
                                #endregion

                                #region //量測數據 MeasureReport

                                if (reportType == "M") { serialNo = serialCode.ToString(); } else { serialNo = QcItemId + '.' + QmmDetailId; }

                                if (measureMachine.IndexOf(MachineNo) == -1)
                                { measureMachine += "(" + MachineNo + ")" + MachineName + " "; }

                                // item 換量測項目的時候，列 + 1，欄回到12
                                if (itemNo != lastItemNo)
                                {
                                    rowIndex++;
                                    colIndex = 12;
                                    serialCode += 1;

                                    //量測項目大於20，要複製工作表，貼下一頁
                                    //要注意的是：換頁的時候，rowIndex要回到13，colIndex要回到12

                                    if (rowIndex % 33 == 0)
                                    {
                                        rowSheetCount += 1;
                                        string copySheetName = DefaultStName + "-" + rowSheetCount.ToString() + ".1";
                                        if (workbook.Worksheets.Any(sheets => sheets.Name.Equals(copySheetName)) == false) //判斷是否存在工作表，不存在則複製
                                        {
                                            firstsheet.CopyTo(copySheetName, workbook.Worksheets.Count);
                                            workbook.Worksheet(copySheetName).Range("B13:U32").Value = "";
                                            workbook.Worksheet(copySheetName).Range("L12:U12").Value = "";
                                            rowIndex = 13;
                                            //measureMachine = "(" + MachineNo + ")" + MachineName + " ";
                                        }
                                    }
                                    itemCount += 1;
                                    if (rowIndex == 14) totalLettering = letteringCount - 1;
                                    letteringCount = 1;
                                }

                                //計算當前sheet位置
                                rowSheetCount = itemCount / 20 == 0 ? itemCount / 20 + 1 : itemCount % 20 == 0 ? itemCount / 20 : itemCount / 20 + 1;
                                calcColSheet = letteringCount / 10 == 0 ? letteringCount / 10 + 1 : letteringCount % 10 == 0 ? letteringCount / 10 : letteringCount / 10 + 1;
                                CurrentSheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + calcColSheet.ToString();

                                //欄位大於21，複製分頁，一頁只放十個pcs的數據
                                //要注意的是：換頁的時候，colIndex要回到12，rowIndex不動
                                if (colIndex % 22 == 0)
                                {
                                    colIndex = 12;
                                    string copySheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + calcColSheet.ToString();
                                    if (workbook.Worksheets.Any(sheets => sheets.Name.Equals(copySheetName)) == false) //判斷是否存在工作表，不存在則複製
                                    {
                                        firstsheet.CopyTo(copySheetName, workbook.Worksheets.Count);
                                        workbook.Worksheet(copySheetName).Range("B13:U32").Value = "";
                                        workbook.Worksheet(copySheetName).Range("L12:U12").Value = "";
                                        colSheetCount += 1;
                                    }
                                }

                                #region //第一筆資料，先放表頭
                                if (rowIndex == 13 && colIndex == 12)
                                {
                                    if (excelFileName == "")
                                    {
                                        if (reportType == "M") { excelFileName = mtlItemName + "_出貨報告_" + DocDate.ToString("yyyyMMdd"); } else { excelFileName = mtlItemName + "_出貨檢量測報告_" + DocDate.ToString("yyyyMMdd"); }
                                    }
                                    workbook.Worksheet(CurrentSheetName).Cell(5, 4).Value = CusShortName; //客戶名稱
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 4).Value = MweErpNo; //來源單號
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 7).Value = QcClassName; //產品類別
                                    workbook.Worksheet(CurrentSheetName).Cell(5, 11).Value = CustomerMtlItemNo; //客戶料號
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 11).Value = mtlItemName; //品名
                                    workbook.Worksheet(CurrentSheetName).Cell(5, 19).Value = DocDate.ToString("yyyy/MM/dd"); //出貨日期
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 19).Value = abnormalNo; //品異單號
                                    workbook.Worksheet(CurrentSheetName).Cell(3, 18).Value = Surveyor; //量測人員
                                    workbook.Worksheet(CurrentSheetName).Cell(3, 21).Value = ""; //審核 or 確認人員
                                }
                                #endregion

                                //放值的時候，要注意分頁的順序
                                #region //換列的時候，先放列的標頭
                                if (colIndex == 12)
                                {
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 2).Value = serialNo; //序號
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 3).Value = BallMark; //球標
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 4).Value = QcItemDesc; //量測項目
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 5).Value = DesignValue; //設計值
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 6).Value = UpperTolerance; //上公差
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 7).Value = LowerTolerance; //下公差
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 8).Value = Unit; //單位
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 9).Value = MachineNo; //量測儀器
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 11).Value = ZAxis; //Z軸
                                    workbook.Worksheet(CurrentSheetName).Cell(34, 4).Value = measureMachine; //量測機台對照欄位
                                }
                                #endregion
                                if (rowIndex == 13)
                                {
                                    pcsNumberList.Add(pcsNumber);
                                    if (inputType == "LotNumber")
                                        workbook.Worksheet(CurrentSheetName).Cell(12, colIndex).Value = pcsNumber + "-" + makeCount; //刻號 or 穴號 or 批號 or 條碼
                                    else
                                        workbook.Worksheet(CurrentSheetName).Cell(12, colIndex).Value = pcsNumber; //刻號 or 穴號 or 批號 or 條碼
                                }
                                workbook.Worksheet(CurrentSheetName).Cell(rowIndex, colIndex).Value = MeasureValue; //量測數據
                                #endregion

                                #region //偏芯表格 Decenter B01k101、B01k301、B01k601、B01k701
                                Regex DecenterReg = new Regex(@"^[A-Z][0-9][0-9][k][0-9][0-9][0-9]");
                                int DecRowIndex = 5;
                                int DecColIndex = 6;
                                int decXRowIndex = 5;
                                int decXColIndex = 3;
                                int decYRowIndex = 5;
                                int decYColIndex = 4;
                                int absColIndex = 5;
                                int absRowIndex = 5;
                                int site = 0;

                                if (DecenterReg.IsMatch(qcItemNo))
                                {
                                    string subitemNo = qcItemNo.Substring(3, 2);
                                    //R1、R2不同列
                                    if (qcItemNo.Substring(6, 1) == "1") site = 0; else if (qcItemNo.Substring(6, 1) == "2") site = 1;
                                    
                                    switch (subitemNo)
                                    {
                                        case "k0": //TILT
                                            //workbook.Worksheet("Decenter").Cell(DecRowIndex, DecColIndex).Value = MeasureValue;
                                            break;
                                        case "k1": //Decenter
                                            if (letteringCount % 4 == 0) //塞滿四個換列
                                            {
                                                DecRowIndex += 16;
                                                DecColIndex = 6;
                                            }
                                            workbook.Worksheet("Decenter").Cell(DecRowIndex, DecColIndex).Value = MeasureValue;
                                            DecColIndex += 6;
                                            break;
                                        case "k3": //偏芯 absolute
                                            if (letteringCount % 4 == 0) //塞滿四個換列
                                            {
                                                absRowIndex += 16;
                                                absColIndex = 5;
                                            }
                                            workbook.Worksheet("Decenter").Cell(absRowIndex + site, absColIndex).Value = MeasureValue;
                                            absColIndex += 6;
                                            break;
                                        case "k6": //偏芯 X 座標
                                            if (letteringCount % 4 == 0) //塞滿四個換列
                                            {
                                                decXRowIndex += 16;
                                                decXColIndex = 3;
                                            }
                                            workbook.Worksheet("Decenter").Cell(decXRowIndex + site, decXColIndex).Value = MeasureValue;
                                            decXColIndex += 6;
                                            break;
                                        case "k7": //偏芯 Y 座標
                                            if (letteringCount % 4 == 0) //塞滿四個換列
                                            {
                                                decYRowIndex += 16;
                                                decYColIndex = 4;
                                            }
                                            workbook.Worksheet("Decenter").Cell(decYRowIndex + site, decYColIndex).Value = MeasureValue;
                                            decYColIndex += 6;
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                #endregion

                                #region //品質工程表***分光數據
                                Regex PenetranceReg = new Regex(@"^[A-Z][0-9][0-9][s][0-9][0-9][0-9]");
                                if (DecenterReg.IsMatch(qcItemNo))
                                {
                                    //取得允收標準及上傳的點資料檔案
                                    string pointRequest = "";
                                    JObject pointjsonResponse = new JObject();

                                    pointRequest = qcReportDA.GetMeasurementPointData(mtlItemId, MoId, QcRecordId);
                                    pointjsonResponse = JObject.Parse(pointRequest);
                                    if (pointjsonResponse["status"].ToString() == "success")
                                    {
                                        if (pointjsonResponse["data"].ToString() != "[]")
                                        {
                                            foreach (var item3 in pointjsonResponse["data"])
                                            {
                                                List<string> result = new List<string>();
                                                result.Add(item3[result].ToString());
                                                if (colIndex == 12)
                                                {
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 2).Value = serialNo; //序號
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 3).Value = BallMark; //球標
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 4).Value = QcItemDesc; //量測項目
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 5).Value = DesignValue; //設計值
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 6).Value = UpperTolerance; //上公差
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 7).Value = LowerTolerance; //下公差
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 8).Value = Unit; //單位
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 9).Value = MachineNo; //量測儀器
                                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, 11).Value = ZAxis; //Z軸
                                                }
                                                workbook.Worksheet(CurrentSheetName).Cell(rowIndex + 1, colIndex).Value = result[0]; //量測數據
                                            }
                                        }
                                    }
                                }

                                colIndex += 1;
                                lastItemNo = itemNo;
                                letteringCount += 1;
                            }
                            #endregion

                            #region //放圖片 SurfaceAccuracy
                            if (reportType == "M")
                            {
                                string dataFileRequest = qcReportDA.GetMeasurementDataFile(QcRecordId, inputType, pcsNumberList);
                                if (JObject.Parse(dataFileRequest)["status"].ToString() == "success")
                                {
                                    dynamic[] file = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataFileRequest)["data"].ToString());

                                    if (file.Count() > 0)
                                    {
                                        int picCount = 1;
                                        int picSheetCount = 1;
                                        int imageRow = 3;
                                        int imageCol = 5;
                                        string CurrentImageSheet = DefaultpicStName + "-1";

                                        foreach (var item in file)
                                        {
                                            string PhysicalPath = item.PhysicalPath;
                                            string EffectiveDiameter = item.EffectiveDiameter;
                                            string Cavity = item.Cavity;
                                            string LotNumber = item.LotNumber;
                                            string Lettering = item.ItemSeq;

                                            string extension = Path.GetExtension(PhysicalPath);

                                            if (extension == ".jpg" && extension == ".png" && extension == ".JPG" && extension == ".PNG")
                                            {
                                                if (picCount > 21)
                                                {
                                                    string copyImageSheet = DefaultpicStName + "-" + (picSheetCount + 1).ToString();
                                                    if (workbook.Worksheets.Any(sheet => sheet.Name.Equals(copyImageSheet)) == false) //判斷是否存在工作表，不存在則複製
                                                    {
                                                        imageSheet.CopyTo(copyImageSheet, workbook.Worksheets.Count + 1);
                                                        var pic = workbook.Worksheet(copyImageSheet).Pictures;
                                                        //複製工作表後，刪除所有圖片
                                                        foreach (var images in pic)
                                                        {
                                                            images.Delete();
                                                        }
                                                    }
                                                    picCount = 1;
                                                    picSheetCount += 1;
                                                    imageCol = 5;
                                                    imageRow = 3;
                                                }

                                                if (picCount > 10)
                                                {
                                                    imageRow = 10;
                                                    imageCol = 5;
                                                }
                                                //先設定R1、R2的位置
                                                int imageRowSite = imageRow;
                                                string picName = "";

                                                if (inputType == "Cavity")
                                                {
                                                    if (EffectiveDiameter.IndexOf("R2") != -1) imageRowSite += 3;
                                                    picName = EffectiveDiameter + '-' + Cavity;
                                                }
                                                else if (inputType == "Lettering")
                                                {
                                                    picName = Lettering;
                                                }
                                                else if (inputType == "LotNumber")
                                                {
                                                    picName = item.LotNumber; //批號
                                                }
                                                else if (inputType == "BarcodeId")
                                                {
                                                    picName = item.BarcodeId; //條碼
                                                }

                                                CurrentImageSheet = DefaultpicStName + "-" + (picSheetCount).ToString();

                                                workbook.Worksheet(CurrentImageSheet).Cell(imageRowSite, imageCol).Value = picName; //檔名
                                                workbook.Worksheet(CurrentImageSheet).Cell(imageRowSite, imageCol).Style.NumberFormat.Format = "General";
                                                int cellH = 350;
                                                int cellW = 300;

                                                using (FileStream fileStream = new FileStream(PhysicalPath, FileMode.Open, FileAccess.Read))
                                                {
                                                    using (MemoryStream imageStream = new MemoryStream())
                                                    {
                                                        fileStream.CopyTo(imageStream);
                                                        imageStream.Position = 0;
                                                        var image = workbook.Worksheet(CurrentImageSheet).AddPicture(imageStream);
                                                        image.Name = picName;
                                                        image.MoveTo(imageSheet.Cell(imageRowSite + 2, imageCol)).WithSize(cellW, cellH);
                                                    }
                                                }
                                                imageCol += 4;
                                                picCount += 1;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //設定
                            #region //自適應欄寬
                            //firstsheet.Columns().AdjustToContents();
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
                            //firstsheet.SheetView.FreezeRows(2);
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
                            //firstsheet.Column(11).Width = 50;
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
                                                        
                            #endregion

                            #region //EXCEL匯出
                            using (MemoryStream output = new MemoryStream())
                            {
                                workbook.SaveAs(output);
                                //excelFile = Encoding.Convert(Encoding.GetEncoding(encoding), Encoding.Default, output.ToArray());
                                excelFile = output.ToArray();
                            }
                            #endregion

                            workbook.Dispose();
                            GC.Collect();
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
                        throw new SystemException("找不到量測資料!!");
                    }
                }
                else
                {
                    throw new SystemException("找不到量測資料!!");
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
        }
        #endregion

        #region //DeliveryReportOutputForExcelSU 出貨報告產出excel 舜宇模仁 -- WuTc -- 2024-06-11
        public void DeliveryReportOutputForExcelSU(string DataId = "", string SearchDate = "", string MtlItemNo = "", string MtlItemName = "", string emptyFileName = "", string encoding = "", string reportType = "")
        {
            try
            {
                #region //Request
                qcReportDA = new QcReportDA();
                if (reportType == "M")
                {
                    dataRequest = qcReportDA.GetDeliveryMeasurementDataModify(DataId, SearchDate, MtlItemNo, MtlItemName, "", -1, -1); //修改後的報告資料
                }
                else
                {
                    dataRequest = qcReportDA.GetDeliveryMeasurementData(DataId, SearchDate, MtlItemNo, MtlItemName, "", -1, -1);
                }
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    if (data.Count() > 0)
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
                        dateStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.TopBorderColor = XLColor.Black;
                        dateStyle.Border.BottomBorderColor = XLColor.Black;
                        dateStyle.Border.LeftBorderColor = XLColor.Black;
                        dateStyle.Border.RightBorderColor = XLColor.Black;
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
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
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
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.DateFormat.Format = "$#,##0.00";
                        #endregion

                        #region //customizedStyle
                        List<XLColor> xLColors = new List<XLColor>
                    {
                        XLColor.TeaRoseRose,
                        XLColor.PeachOrange,
                        XLColor.Flavescent,
                        XLColor.MediumSpringBud,
                        XLColor.EtonBlue,
                        XLColor.MossGreen,
                        XLColor.LightBlue,
                        XLColor.WildBlueYonder,
                        XLColor.Wisteria,
                        XLColor.PinkPearl
                    };

                        List<XLBorderStyleValues> xLBorderSizes = new List<XLBorderStyleValues>
                    {
                        XLBorderStyleValues.Thin,
                        XLBorderStyleValues.Thick
                    };

                        List<XLColor> xLBorderColors = new List<XLColor>
                    {
                        XLColor.Black,
                        XLColor.CornflowerBlue
                    };
                        #endregion
                        #endregion

                        #region //參數初始化

                        byte[] excelFile;
                        string excelFileName = "";
                        string DefaultStName = "MeasureReport";
                        string DefaultpicStName = "SurfaceAccuracy";

                        int colIndex = 10;
                        int rowIndex = 3;
                        #endregion

                        #region //EXCEL
                        using (var workbook = new XLWorkbook(Server.MapPath("~/PdfTemplate/MES/QcReport/Shipment/" + emptyFileName)))
                        {
                            IXLWorksheet firstsheet = workbook.Worksheet(DefaultStName + "-1.1");
                            IXLWorksheet imageSheet = workbook.Worksheet(DefaultpicStName + "-1");

                            #region //圖片
                            //var imagePath = Server.MapPath("~/Content/images/logo.png");
                            //var image = firstsheet.AddPicture(imagePath).MoveTo(firstsheet.Cell(1, 2)).Scale(0.07);
                            //firstsheet.Row(rowIndex).Height = 40;
                            #endregion

                            #region //HEADER

                            #endregion

                            #region //BODY
                            string lastItemNo = "";
                            int serialCode = 0;
                            int rowSheetCount = 1;
                            int colSheetCount = 2;
                            string CurrentSheetName = DefaultStName + "-1.1";
                            string measureMachine = "量測儀器：";
                            int itemCount = 1;
                            int letteringCount = 0;
                            int totalLettering = 0;
                            int calcColSheet = -1;
                            int insertRowCount = 0;
                            string QcRecordId = "";
                            string inputType = "";
                            List<string> pcsNumberList = new List<string> { };
                            string QmdIdList = "";
                            string mtlItemId = "";
                            string MoId = "";

                            foreach (var item in data)
                            {
                                #region //宣告
                                mtlItemId = item.MtlItemId;
                                string QmdId = item.QmdId;
                                string itemNo = item.QcItemNo + item.MachineNumber;
                                string BallMark = item.BallMark;
                                string QcItemDesc = item.QcItemDesc;
                                string UpperTolerance = item.UpperTolerance;
                                string LowerTolerance = item.LowerTolerance;
                                string Unit = item.Unit;
                                string MachineNo = item.MachineNo;
                                string MachineName = item.MachineName;
                                string ZAxis = item.ZAxis;
                                string Surveyor = "";
                                string CusShortName = item.CustomerShortName;
                                string MweErpNo = item.MweErpNo;
                                string QcClassName = item.QcClassName;
                                string CustomerMtlItemNo = item.CustomerMtlItemNo;
                                string CustomerDwgNo = item.CustomerDwgNo;
                                string mtlItemName = item.MtlItemName;
                                MtlItemNo = item.MtlItemNo;
                                string abnormalNo = item.AbnormalqualityNo;
                                string pcsNumber = "";
                                string DesignValue = item.DesignValue;
                                string QcItemId = item.QcItemId;
                                string QmmDetailId = item.QmmDetailId;
                                string BarcodeId = item.QcBarcodeId;
                                MoId = item.MoId;
                                string MachineNumber = item.MachineNumber;
                                string Confirmer = "";
                                inputType = item.InputType;
                                QcRecordId = item.QcRecordId;
                                //int totalQuantity = item.Quantity;

                                string MeasureValue = "";
                                DateTime DocDate;
                                string serialNo = "";
                                string machineNo = "";
                                string lettering = "";

                                if (reportType == "M")
                                {
                                    machineNo = MachineNo;
                                    lettering = item.LetteringSeq; //刻號
                                    MeasureValue = item.ModifyValue;
                                    DocDate = item.CreateDate;
                                    Surveyor = item.Surveyor;
                                    Confirmer = item.Confirmer;
                                }
                                else
                                {
                                    if (inputType == "Lettering")
                                    {
                                        pcsNumber = item.ItemSeq; //刻號
                                    }
                                    else if (inputType == "Cavity")
                                    {
                                        pcsNumber = item.Cavity; //穴號
                                    }
                                    else if (inputType == "LotNumber")
                                    {
                                        pcsNumber = item.LotNumber; //批號
                                    }
                                    else if (inputType == "BarcodeId")
                                    {
                                        pcsNumber = item.BarcodeId; //條碼
                                    }

                                    machineNo = QmmDetailId;
                                    lettering = pcsNumber + " (" + BarcodeId + ")";
                                    MeasureValue = item.MeasureValue;
                                    DocDate = item.DocDate;
                                    Surveyor = item.QcUserName;
                                }
                                #endregion

                                #region //量測數據 MeasureReport
                                // item 換量測項目的時候，列 + 1，欄回到10
                                if (itemNo != lastItemNo)
                                {
                                    rowIndex++;
                                    colIndex = 10;
                                    serialCode += 1;
                                    if (reportType == "M")
                                    {
                                        serialNo = serialCode.ToString();
                                    }
                                    else
                                    {
                                        serialNo = QcItemId + '.' + QmmDetailId;
                                    }

                                    //量測項目大於100，要複製列
                                    if (rowIndex % 50 == 0)
                                    {
                                        workbook.Worksheet(CurrentSheetName).Row(rowIndex).InsertRowsBelow(1);
                                        var copyRange = workbook.Worksheet(CurrentSheetName).Range("A" + rowIndex + 1 + ":AH" + rowIndex + 1);
                                        workbook.Worksheet(CurrentSheetName).Range("A4:AH4").CopyTo(copyRange);
                                        workbook.Worksheet(CurrentSheetName).Range("A4:G4").Value = "";
                                        workbook.Worksheet(CurrentSheetName).Range("J4:AD4").Value = "";
                                        insertRowCount += 1;
                                    }

                                    //量測機台對照欄位
                                    if (measureMachine.IndexOf(MachineNo) == -1)
                                    {
                                        measureMachine += "(" + MachineNo + ")" + MachineName + " ";
                                        workbook.Worksheet(CurrentSheetName).Cell(51 + insertRowCount, 2).Value = measureMachine;
                                    }
                                    itemCount += 1;
                                    if (rowIndex == 5) totalLettering = letteringCount - 1;
                                    letteringCount = 1;
                                }

                                //計算當前sheet位置
                                calcColSheet = letteringCount / 21 == 0 ? letteringCount / 21 + 1 : letteringCount % 21 == 0 ? letteringCount / 21 : letteringCount / 21 + 1;
                                CurrentSheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + calcColSheet.ToString();

                                //欄位大於21，複製分頁，一頁只放十個pcs的數據
                                //要注意的是：換頁的時候，colIndex要回到12，rowIndex不動
                                if (colIndex % 31 == 0)
                                {
                                    colIndex = 10;
                                    string copySheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + calcColSheet.ToString();
                                    if (workbook.Worksheets.Any(sheets => sheets.Name.Equals(copySheetName)) == false) //判斷是否存在工作表，不存在則複製
                                    {
                                        firstsheet.CopyTo(copySheetName, workbook.Worksheets.Count);
                                        workbook.Worksheet(copySheetName).Range("A4:G100").Value = "";
                                        workbook.Worksheet(copySheetName).Range("J3:AD100").Value = "";
                                        colSheetCount += 1;
                                    }
                                }

                                //第一筆資料，先放表頭
                                if (rowIndex == 4)
                                {
                                    if (excelFileName == "")
                                    {
                                        if (reportType == "M")
                                            excelFileName = mtlItemName + "_出貨報告_" + DocDate.ToString("yyyyMMdd");
                                        else
                                            excelFileName = mtlItemName + "_出貨檢量測報告_" + DocDate.ToString("yyyyMMdd");
                                    }
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 3).Value = CusShortName; //客戶名稱
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 9).Value = CustomerMtlItemNo; //客戶料號
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 13).Value = ""; //版本號
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 20).Value = Confirmer; //審核 or 確認人員
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 23).Value = Surveyor; //量測人員
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 26).Value = "白"; //班次
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 28).Value = CustomerDwgNo; //圖面版號
                                    workbook.Worksheet(CurrentSheetName).Cell(2, 31).Value = DocDate.ToString("yyyy/MM/dd"); //出貨日期
                                }

                                //放值的時候，要注意分頁的順序
                                #region //換列的時候，先放列的標頭
                                if (colIndex == 10)
                                {
                                    workbook.Worksheet(CurrentSheetName).Cell(1, 1).Value = MoId + " (" + QcRecordId + ")";
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 1).Value = serialNo; //序號 --> serialCode or QcItemId
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 3).Value = machineNo; //量測儀器 --> MachineName or  QmmDetailId
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 2).Value = QcItemDesc; //量測項目
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 4).Value = DesignValue; //設計值
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 5).Value = UpperTolerance; //上公差
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 6).Value = LowerTolerance; //下公差
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 7).Value = Unit; //單位
                                }
                                #endregion
                                if (rowIndex == 4)
                                {
                                    pcsNumberList.Add(lettering);
                                    workbook.Worksheet(CurrentSheetName).Cell(3, colIndex).Value = lettering; //刻號 or 穴號 or 批號 or 條碼
                                }

                                //找到同穴號/刻號的欄，再放值
                                for (int i = 10; i < 31; i++)
                                {
                                    var cellLettering = workbook.Worksheet(CurrentSheetName).Cell(3, i).Value.ToString();
                                    if (cellLettering == lettering)
                                    {
                                        if (QmdIdList != "") QmdIdList += ",";
                                        QmdIdList += QmdId;
                                        workbook.Worksheet(CurrentSheetName).Cell(rowIndex, i).Value = MeasureValue; //量測數據
                                        break;
                                    }
                                }
                                #endregion

                                colIndex += 1;
                                lastItemNo = itemNo;
                                letteringCount += 1;
                            }

                            #region //放圖片 SurfaceAccuracy
                            if (reportType == "M")
                            {
                                string dataFileRequest = qcReportDA.GetMeasurementDataFile(QcRecordId, inputType, pcsNumberList);
                                if (JObject.Parse(dataFileRequest)["status"].ToString() == "success")
                                {
                                    dynamic[] file = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataFileRequest)["data"].ToString());

                                    if (file.Count() > 0)
                                    {
                                        string CurrentImageSheet = DefaultpicStName + "-1";
                                        int picCount = 1;
                                        int picSheetCount = 1;
                                        int imageRow = 3;
                                        int imageCol = 2;

                                        foreach (var item in file)
                                        {
                                            string PhysicalPath = item.PhysicalPath;
                                            string EffectiveDiameter = item.EffectiveDiameter;
                                            string Cavity = item.Cavity;
                                            string LotNumber = item.LotNumber;
                                            string Lettering = item.ItemSeq;

                                            if (picCount > 41)
                                            {
                                                string copyImageSheet = DefaultpicStName + "-" + (picSheetCount + 1).ToString();
                                                if (workbook.Worksheets.Any(sheet => sheet.Name.Equals(copyImageSheet)) == false) //判斷是否存在工作表，不存在則複製
                                                {
                                                    imageSheet.CopyTo(copyImageSheet, workbook.Worksheets.Count + 1);
                                                    var pic = workbook.Worksheet(copyImageSheet).Pictures;
                                                    //複製工作表後，刪除所有圖片
                                                    foreach (var images in pic)
                                                    {
                                                        images.Delete();
                                                    }
                                                }
                                                picCount = 1;
                                                picSheetCount += 1;
                                                imageCol = 2;
                                                imageRow = 3;
                                            }
                                            if (PhysicalPath != "")
                                            {
                                                if (picCount > 10)
                                                {
                                                    imageRow += 3;
                                                    imageCol = 2;
                                                }
                                                else if (picCount > 20)
                                                {
                                                    imageRow += 4;
                                                    imageCol = 2;
                                                }
                                                else if (picCount > 30)
                                                {
                                                    imageRow += 3;
                                                    imageCol = 2;
                                                }
                                                //先設定R1、R2的位置
                                                string picName = "";

                                                if (inputType == "Cavity")
                                                {
                                                    picName = Cavity;
                                                }
                                                else if (inputType == "Lettering")
                                                {
                                                    picName = Lettering;
                                                }
                                                else if (inputType == "LotNumber")
                                                {
                                                    picName = item.LotNumber; //批號
                                                }
                                                else if (inputType == "BarcodeId")
                                                {
                                                    picName = item.BarcodeId; //條碼
                                                }

                                                CurrentImageSheet = DefaultpicStName + "-" + (picSheetCount).ToString();

                                                workbook.Worksheet(CurrentImageSheet).Cell(imageRow, imageCol).Value = picName; //檔名
                                                workbook.Worksheet(CurrentImageSheet).Cell(imageRow, imageCol).Style.NumberFormat.Format = "General";
                                                int cellH = 400;
                                                int cellW = 330;
                                                using (FileStream fileStream = new FileStream(PhysicalPath, FileMode.Open, FileAccess.Read))
                                                {
                                                    using (MemoryStream imageStream = new MemoryStream())
                                                    {
                                                        fileStream.CopyTo(imageStream);
                                                        imageStream.Position = 0;
                                                        var image = workbook.Worksheet(CurrentImageSheet).AddPicture(imageStream);
                                                        image.Name = picName;
                                                        image.MoveTo(imageSheet.Cell(imageRow + 2, imageCol)).WithSize(cellW, cellH);
                                                    }
                                                }
                                                imageCol += 4;
                                                picCount += 1;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //把下方多出來的空列，隱藏
                            for (int s = 1; s < workbook.Worksheets.Count(); s++)
                            {
                                IXLWorksheet worksheet = workbook.Worksheet(s);
                                if (worksheet.Name.IndexOf(DefaultStName) != -1)
                                {
                                    int usedRows = worksheet.RowsUsed().Count() - 1;
                                    for (int i = 4; i < usedRows; i++)
                                    {
                                        string serial = worksheet.Cell(i, 1).Value.ToString();
                                        if (serial == "")
                                        {
                                            worksheet.Row(i).Hide();
                                        }
                                    }
                                }
                            }
                            #endregion

                            #endregion

                            #region //設定
                            #region //自適應欄寬
                            //firstsheet.Columns().AdjustToContents();
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
                            //firstsheet.SheetView.FreezeRows(2);
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
                            //firstsheet.Column(11).Width = 50;
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
                            workbook.Dispose();
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
                    else throw new SystemException("找不到量測資料");
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

        #region //UploadModifyMeasureDataForMold 上傳修改後的出貨報告excel -- WuTc -- 2024-08-09
        public void UploadModifyMeasureDataForMold(string MoIdErpNo = "", string BatchNo = "", string FileType = "", string QcRecordFileByNas = "")
        {
            List<UploadModifyDataModel> ModifyDataSetList = new List<UploadModifyDataModel>();
            try
            {
                WebApiLoginCheck("QcDeliveryReportModify", "add");

                #region //共夾路徑上傳
                JObject dataRequestJson = new JObject();
                JObject DataListJson = new JObject();
                List<string> FileList = new List<string>();
                JObject FileJson = JObject.Parse(QcRecordFileByNas);
                if (FileJson["uploadFileInfo"].Count() <= 0) throw new SystemException("請上傳出貨報告excel檔案!!");
                #endregion

                foreach (var item in FileJson["uploadFileInfo"])
                {
                    string FilePath = item["FilePath"].ToString();
                    string QcRecordId = "";
                    string MoId = "";
                    using (var workbook = new XLWorkbook(FilePath))
                    {
                        //只上傳數據，不上傳圖片
                        string DefaultStName = "MeasureReport";
                        string headerCell = workbook.Worksheet(DefaultStName + "-1.1").Cell(1, 1).Value.ToString();
                        MoId = headerCell.Split('(')[0].Trim();
                        QcRecordId = headerCell.Split('(')[1].Replace(')', ' ').Trim();

                        if (QcRecordId == "") throw new SystemException("找不到原始量測單據，請檢查是否上傳正確檔案！");

                        #region //遍歷資料表，找到放資料的工作頁
                        if (FileType == "SuMode")
                        {
                            //舜宇模仁表格-2024-10-19
                            for (int s = 1; s < workbook.Worksheets.Count(); s++)
                            {
                                IXLWorksheet worksheet = workbook.Worksheet(s);
                                if (workbook.Worksheet(s).Name.IndexOf(DefaultStName) != -1)
                                {
                                    int usedRows = worksheet.RowsUsed().Count();
                                    int usedColumns = worksheet.ColumnsUsed().Count();
                                    string Confirmer = worksheet.Cell(2, 20).Value.ToString();
                                    string Surveyor = worksheet.Cell(2, 23).Value.ToString();

                                    for (int i = 10; i < usedColumns; i++) //遍歷每一欄
                                    {
                                        string letterCell = worksheet.Cell(3, i).Value.ToString();
                                        if (letterCell != "" && letterCell.IndexOf('(') != -1) //刻號不為空，且格式為***(**)
                                        {
                                            string Lettering = letterCell.Split('(')[0].Replace(" ", "");
                                            string BarcodeId = letterCell.Split('(')[1].Replace(')', ' ').Replace(" ", "");

                                            for (int j = 4; j < usedRows; j++)//遍歷每一列
                                            {
                                                string QcItemId = worksheet.Cell(j, 1).Value.ToString();
                                                if (QcItemId != "")
                                                {
                                                    string QcItemDesc = worksheet.Cell(j, 2).Value.ToString();
                                                    string MachineId = worksheet.Cell(j, 3).Value.ToString();
                                                    string DesignValue = worksheet.Cell(j, 4).Value.ToString();
                                                    string Usl = worksheet.Cell(j, 5).Value.ToString();
                                                    string Lsl = worksheet.Cell(j, 6).Value.ToString();
                                                    string Unit = worksheet.Cell(j, 7).Value.ToString();
                                                    string MeasureValue = worksheet.Cell(j, i).Value.ToString();
                                                    string ZAxis = "";
                                                    string BallMark = "";

                                                    //將值拋到da，新增到QcMeasureDataModify
                                                    if (MeasureValue != "")
                                                    {
                                                        UploadModifyDataModel ModifyDataSet = new UploadModifyDataModel()
                                                        {
                                                            //QmdId = QmdId,
                                                            QcRecordId = Convert.ToInt32(QcRecordId),
                                                            MoId = Convert.ToInt32(MoId),
                                                            QcItemId = Convert.ToInt32(QcItemId),
                                                            QcItemDesc = QcItemDesc,
                                                            DesignValue = DesignValue,
                                                            UpperTolerance = Usl,
                                                            LowerTolerance = Lsl,
                                                            ZAxis = ZAxis,
                                                            ModifyValue = MeasureValue,
                                                            BarcodeId = Convert.ToInt32(BarcodeId),
                                                            LetteringSeq = Lettering,
                                                            QmmDetailId = MachineId,
                                                            BallMark = BallMark,
                                                            Unit = Unit,
                                                            Surveyor = Surveyor,
                                                            Confirmer = Confirmer,
                                                        };

                                                        ModifyDataSetList.Add(ModifyDataSet);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (FileType == "Standard")
                        {
                            //新版制式表格
                            for (int s = 1; s < workbook.Worksheets.Count(); s++)
                            {
                                IXLWorksheet worksheet = workbook.Worksheet(s);
                                if (workbook.Worksheet(s).Name.IndexOf(DefaultStName) != -1)
                                {
                                    int usedRows = worksheet.RowsUsed().Count();
                                    int usedColumns = worksheet.ColumnsUsed().Count();
                                    string Confirmer = worksheet.Cell(3, 21).Value.ToString();
                                    string Surveyor = worksheet.Cell(3, 18).Value.ToString();

                                    for (int i = 12; i < 22; i++) //遍歷每一欄，共十欄
                                    {
                                        string letterCell = worksheet.Cell(12, i).Value.ToString();
                                        if (letterCell != "" && letterCell.IndexOf('(') != -1) //刻號不為空，且格式為***(**)
                                        {
                                            string Lettering = letterCell.Split('(')[0].Replace(" ", "");
                                            string BarcodeId = letterCell.Split('(')[1].Replace(')', ' ').Replace(" ", "");

                                            for (int j = 12; j < 33; j++)//遍歷每一列，共二十列
                                            {
                                                string QcItemId = worksheet.Cell(j, 2).Value.ToString().Split('.')[0];
                                                if (QcItemId != "")
                                                {
                                                    string BallMark = worksheet.Cell(j, 3).Value.ToString();
                                                    string QcItemDesc = worksheet.Cell(j, 4).Value.ToString();
                                                    string DesignValue = worksheet.Cell(j, 5).Value.ToString();
                                                    string Usl = worksheet.Cell(j, 6).Value.ToString();
                                                    string Lsl = worksheet.Cell(j, 7).Value.ToString();
                                                    string Unit = worksheet.Cell(j, 8).Value.ToString();
                                                    string MachineId = worksheet.Cell(j, 9).Value.ToString();
                                                    string ZAxis = worksheet.Cell(j, 11).Value.ToString();
                                                    string MeasureValue = worksheet.Cell(j, i).Value.ToString();

                                                    //將值拋到da，新增到QcMeasureDataModify
                                                    if (MeasureValue != "")
                                                    {
                                                        UploadModifyDataModel ModifyDataSet = new UploadModifyDataModel()
                                                        {
                                                            //QmdId = QmdId,
                                                            QcRecordId = Convert.ToInt32(QcRecordId),
                                                            MoId = Convert.ToInt32(MoId),
                                                            QcItemId = Convert.ToInt32(QcItemId),
                                                            QcItemDesc = QcItemDesc,
                                                            DesignValue = DesignValue,
                                                            UpperTolerance = Usl,
                                                            LowerTolerance = Lsl,
                                                            ZAxis = ZAxis,
                                                            ModifyValue = MeasureValue,
                                                            BarcodeId = Convert.ToInt32(BarcodeId),
                                                            LetteringSeq = Lettering,
                                                            QmmDetailId = MachineId,
                                                            BallMark = BallMark,
                                                            Unit = Unit,
                                                            Surveyor = Surveyor,
                                                            Confirmer = Confirmer,
                                                        };

                                                        ModifyDataSetList.Add(ModifyDataSet);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                        workbook.Dispose();

                        #region //INSERT Modify Data
                        dataRequest = qcReportDA.UploadModifyMeasureDataForMold(ModifyDataSetList, QcRecordId);
                        #endregion

                        #region //共夾的路徑存到DB
                        dataRequest = qcReportDA.AddMeasureDataModifyFile(QcRecordId, FilePath, "upload-ModityExcel");
                        #endregion
                    }
                }
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

        #region //GoodsReceiptForExcelStandard 進貨檢報告產出excel 制式表格 -- WuTc -- 2024-10-05
        public void GoodsReceiptForExcelStandard(string QcRecordId = "", string GrDetailId = "", string SearchDate = "", string MtlItemNo = "", string MtlItemName = "")
        {
            try
            {
                #region //Request
                qcReportDA = new QcReportDA();
                dataRequest = qcReportDA.GetGoodsReceiptData(QcRecordId, GrDetailId, SearchDate, MtlItemNo, MtlItemName, "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    if (data.Count() > 0)
                    {
                        string emptyFileName = "Quality_Inspection_Standard_IQC_Report.xlsx";
                        string encoding = "UTF-8";

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
                        dateStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dateStyle.Border.TopBorderColor = XLColor.Black;
                        dateStyle.Border.BottomBorderColor = XLColor.Black;
                        dateStyle.Border.LeftBorderColor = XLColor.Black;
                        dateStyle.Border.RightBorderColor = XLColor.Black;
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
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
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
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.DateFormat.Format = "$#,##0.00";
                        #endregion

                        #region //customizedStyle
                        List<XLColor> xLColors = new List<XLColor>
                    {
                        XLColor.TeaRoseRose,
                        XLColor.PeachOrange,
                        XLColor.Flavescent,
                        XLColor.MediumSpringBud,
                        XLColor.EtonBlue,
                        XLColor.MossGreen,
                        XLColor.LightBlue,
                        XLColor.WildBlueYonder,
                        XLColor.Wisteria,
                        XLColor.PinkPearl
                    };

                        List<XLBorderStyleValues> xLBorderSizes = new List<XLBorderStyleValues>
                    {
                        XLBorderStyleValues.Thin,
                        XLBorderStyleValues.Thick
                    };

                        List<XLColor> xLBorderColors = new List<XLColor>
                    {
                        XLColor.Black,
                        XLColor.CornflowerBlue
                    };
                        #endregion
                        #endregion

                        #region //參數初始化

                        byte[] excelFile;
                        string excelFileName = "";
                        string DefaultStName = "MeasureReport";
                        string DefaultpicStName = "SurfaceAccuracy";

                        int colIndex = 12;
                        int rowIndex = 12;
                        #endregion

                        #region //EXCEL
                        using (var workbook = new XLWorkbook(Server.MapPath("~/PdfTemplate/MES/QcReport/Shipment/" + emptyFileName)))
                        {
                            IXLWorksheet firstsheet = workbook.Worksheet(DefaultStName + "-1.1");
                            IXLWorksheet imageSheet = workbook.Worksheet(DefaultpicStName + "-1");

                            #region //圖片
                            //var imagePath = Server.MapPath("~/Content/images/logo.png");
                            //var image = firstsheet.AddPicture(imagePath).MoveTo(firstsheet.Cell(1, 2)).Scale(0.07);
                            //firstsheet.Row(rowIndex).Height = 40;
                            #endregion

                            #region //HEADER

                            #endregion

                            #region //BODY
                            string lastItemNo = "";
                            int serialCode = 0;
                            int rowSheetCount = 1;
                            int colSheetCount = 2;
                            string CurrentSheetName = DefaultStName + "-1.1";
                            string measureMachine = "";
                            int itemCount = 1;
                            int letteringCount = 0;
                            int totalLettering = 0;
                            int calcColSheet = -1;
                            string inputType = "";
                            List<string> pcsNumberList = new List<string> { };

                            foreach (var item in data)
                            {
                                #region //宣告
                                string itemNo = item.QcItemNo;
                                string BallMark = item.BallMark;
                                string QcItemDesc = item.QcItemDesc;
                                string DesignValue = item.DesignValue;
                                string UpperTolerance = item.UpperTolerance;
                                string LowerTolerance = item.LowerTolerance;
                                string Unit = item.Unit;
                                string MachineNo = item.MachineNo;
                                string MachineName = item.MachineName;
                                string ZAxis = item.ZAxis == 0 ? null : item.ZAxis;
                                string QcUserName = item.QcUserName;
                                string SupplierName = item.SupplierName;
                                string GrFullErpNo = item.GrFullErpNo;
                                string QcClassName = item.QcClassName;
                                string GrMtlItemName = item.GrMtlItemName;
                                MtlItemNo = item.MtlItemNo;
                                DateTime MeasureDate = item.MeasureDate;
                                string abnormalNo = item.AbnormalqualityNo;
                                string MeasureValue = item.MeasureValue;
                                string pcsNumber = "";
                                int totalQuantity = item.Quantity;
                                inputType = item.InputType;

                                if (inputType == "Lettering")
                                {
                                    pcsNumber = item.ItemSeq; //刻號
                                }
                                else if (inputType == "Cavity")
                                {
                                    pcsNumber = item.Cavity; //穴號
                                }
                                else if (inputType == "LotNumber")
                                {
                                    pcsNumber = item.LotNumber; //批號
                                }
                                else if (inputType == "BarcodeId")
                                {
                                    pcsNumber = item.LotNumber; //批號
                                }

                                #endregion

                                #region //量測數據 MeasureReport
                                // item 換量測項目的時候，列 + 1，欄回到12
                                if (itemNo != lastItemNo)
                                {
                                    rowIndex++;
                                    colIndex = 12;
                                    serialCode += 1;

                                    if (measureMachine.IndexOf(MachineNo) == -1)
                                    {
                                        measureMachine += "(" + MachineNo + ")" + MachineName + " ";
                                    }

                                    //量測項目大於20，要複製工作表，貼下一頁
                                    //要注意的是：換頁的時候，rowIndex要回到13，colIndex要回到12
                                    if (rowIndex % 33 == 0)
                                    {
                                        if (itemCount % 21 == 0) rowSheetCount += 1;
                                        CurrentSheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + (calcColSheet - 1).ToString();
                                        if (workbook.Worksheets.Any(sheets => sheets.Name.Equals(CurrentSheetName)) == false) //判斷是否存在工作表，不存在則複製
                                        {
                                            firstsheet.CopyTo(CurrentSheetName, workbook.Worksheets.Count);
                                            workbook.Worksheet(CurrentSheetName).Range("B13:U32").Value = "";
                                            workbook.Worksheet(CurrentSheetName).Range("L12:U12").Value = "";
                                            rowIndex = 13;
                                            //measureMachine = "(" + MachineNo + ")" + MachineName + " ";
                                        }
                                    }
                                    itemCount += 1;
                                    if (rowIndex == 14) totalLettering = letteringCount - 1;
                                    letteringCount = 1;
                                }

                                //計算當前sheet位置
                                calcColSheet = letteringCount / 10 == 0 ? letteringCount / 10 + 1 : letteringCount % 10 == 0 ? letteringCount / 10 : letteringCount / 10 + 1;
                                CurrentSheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + calcColSheet.ToString();

                                //欄位大於21，複製分頁，一頁只放十個pcs的數據
                                //要注意的是：換頁的時候，colIndex要回到12，rowIndex不動
                                if (colIndex % 22 == 0)
                                {
                                    colIndex = 12;
                                    string copySheetName = DefaultStName + "-" + rowSheetCount.ToString() + "." + calcColSheet.ToString();
                                    if (workbook.Worksheets.Any(sheets => sheets.Name.Equals(copySheetName)) == false) //判斷是否存在工作表，不存在則複製
                                    {
                                        firstsheet.CopyTo(copySheetName, workbook.Worksheets.Count);
                                        workbook.Worksheet(copySheetName).Range("B13:U32").Value = "";
                                        workbook.Worksheet(copySheetName).Range("L12:U12").Value = "";
                                        colSheetCount += 1;
                                    }
                                }

                                //第一筆資料，先放表頭
                                if (rowIndex == 13 && colIndex == 12)
                                {
                                    excelFileName = GrMtlItemName + "_進料檢驗報告_" + MeasureDate.ToString("yyyyMMdd");
                                    workbook.Worksheet(CurrentSheetName).Cell(5, 4).Value = SupplierName; //供應商名稱
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 4).Value = GrFullErpNo; //來源單號
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 7).Value = QcClassName; //產品類別
                                    workbook.Worksheet(CurrentSheetName).Cell(5, 11).Value = MtlItemNo; //產品料號
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 11).Value = GrMtlItemName; //供應商品名
                                    workbook.Worksheet(CurrentSheetName).Cell(5, 19).Value = MeasureDate.ToString("yyyy/MM/dd"); //檢驗日期
                                    workbook.Worksheet(CurrentSheetName).Cell(6, 19).Value = abnormalNo; //品異單號
                                    workbook.Worksheet(CurrentSheetName).Cell(3, 18).Value = QcUserName; //量測人員
                                    workbook.Worksheet(CurrentSheetName).Cell(3, 21).Value = ""; //審核 or 確認人員
                                }

                                //放值的時候，要注意分頁的順序
                                #region //換列的時候，先放列的標頭
                                if (colIndex == 12)
                                {
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 2).Value = serialCode; //序號
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 3).Value = BallMark; //球標
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 4).Value = QcItemDesc; //量測項目
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 5).Value = DesignValue; //設計值
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 6).Value = UpperTolerance; //上公差
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 7).Value = LowerTolerance; //下公差
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 8).Value = Unit; //單位
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 9).Value = MachineNo; //量測儀器
                                    workbook.Worksheet(CurrentSheetName).Cell(rowIndex, 11).Value = ZAxis; //Z軸
                                    workbook.Worksheet(CurrentSheetName).Cell(34, 4).Value = measureMachine; //量測機台對照欄位
                                }
                                #endregion
                                if (rowIndex == 13)
                                {
                                    pcsNumberList.Add(pcsNumber);
                                    workbook.Worksheet(CurrentSheetName).Cell(12, colIndex).Value = pcsNumber; //刻號 or 穴號 or 批號 or 條碼
                                }
                                workbook.Worksheet(CurrentSheetName).Cell(rowIndex, colIndex).Value = MeasureValue; //量測數據
                                #endregion

                                colIndex += 1;
                                lastItemNo = itemNo;
                                letteringCount += 1;
                            }
                            #endregion

                            #region //放圖片 SurfaceAccuracy
                            string dataFileRequest = qcReportDA.GetMeasurementDataFile(QcRecordId, inputType, pcsNumberList);
                            if (JObject.Parse(dataFileRequest)["status"].ToString() == "success")
                            {
                                dynamic[] file = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataFileRequest)["data"].ToString());

                                if (file.Count() > 0)
                                {
                                    int picCount = 1;
                                    int picSheetCount = 1;
                                    int imageRow = 3;
                                    int imageCol = 5;
                                    string CurrentImageSheet = DefaultpicStName + "-1";

                                    foreach (var item in file)
                                    {
                                        string PhysicalPath = item.PhysicalPath;
                                        string EffectiveDiameter = item.EffectiveDiameter;
                                        string Cavity = item.Cavity;
                                        string LotNumber = item.LotNumber;
                                        string Lettering = item.ItemSeq;

                                        if (picCount > 21)
                                        {
                                            string copyImageSheet = DefaultpicStName + "-" + (picSheetCount + 1).ToString();
                                            if (workbook.Worksheets.Any(sheet => sheet.Name.Equals(copyImageSheet)) == false) //判斷是否存在工作表，不存在則複製
                                            {
                                                imageSheet.CopyTo(copyImageSheet, workbook.Worksheets.Count + 1);
                                                var pic = workbook.Worksheet(copyImageSheet).Pictures;
                                                //複製工作表後，刪除所有圖片
                                                foreach (var images in pic)
                                                {
                                                    images.Delete();
                                                }
                                            }
                                            picCount = 1;
                                            picSheetCount += 1;
                                            imageCol = 5;
                                            imageRow = 3;
                                        }
                                        if (PhysicalPath != "")
                                        {
                                            if (picCount > 10)
                                            {
                                                imageRow = 10;
                                                imageCol = 5;
                                            }
                                            //先設定R1、R2的位置
                                            int imageRowSite = imageRow;
                                            string picName = "";

                                            if (inputType == "Cavity")
                                            {
                                                if (EffectiveDiameter.IndexOf("R2") != -1) imageRowSite += 3;
                                                picName = EffectiveDiameter + '-' + Cavity;
                                            }
                                            else if (inputType == "Lettering")
                                            {
                                                picName = Lettering;
                                            }
                                            else if (inputType == "LotNumber")
                                            {
                                                picName = item.LotNumber; //批號
                                            }
                                            else if (inputType == "BarcodeId")
                                            {
                                                picName = item.BarcodeId; //條碼
                                            }

                                            CurrentImageSheet = DefaultpicStName + "-" + (picSheetCount).ToString();

                                            workbook.Worksheet(CurrentImageSheet).Cell(imageRowSite, imageCol).Value = picName; //檔名
                                            workbook.Worksheet(CurrentImageSheet).Cell(imageRowSite, imageCol).Style.NumberFormat.Format = "General";
                                            int cellH = 350;
                                            int cellW = 300;
                                            using (FileStream fileStream = new FileStream(PhysicalPath, FileMode.Open, FileAccess.Read))
                                            {
                                                using (MemoryStream imageStream = new MemoryStream())
                                                {
                                                    fileStream.CopyTo(imageStream);
                                                    imageStream.Position = 0;
                                                    var image = workbook.Worksheet(CurrentImageSheet).AddPicture(imageStream);
                                                    image.Name = picName;
                                                    image.MoveTo(imageSheet.Cell(imageRowSite + 2, imageCol)).WithSize(cellW, cellH);
                                                }
                                            }
                                            imageCol += 4;
                                            picCount += 1;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //設定
                            #region //自適應欄寬
                            //firstsheet.Columns().AdjustToContents();
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
                            //firstsheet.SheetView.FreezeRows(2);
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
                            //firstsheet.Column(11).Width = 50;
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

                            //workbook.SaveAs(Server.MapPath("~/PdfTemplate/MES/QcReport/Shipment/" + excelFileName + ".xlsx"));

                            #region //EXCEL匯出
                            using (MemoryStream output = new MemoryStream())
                            {
                                workbook.SaveAs(output);
                                excelFile = Encoding.Convert(Encoding.GetEncoding(encoding), Encoding.Default, output.ToArray());
                            }
                            #endregion
                            workbook.Dispose();
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
                        throw new SystemException("找不到量測資料!!");
                    }
                }
                else
                {
                    throw new SystemException("找不到量測資料!!");
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
                Response.Write(jsonResponse.ToString());
            }
        }
        #endregion
        #endregion
        #endregion
    }
}