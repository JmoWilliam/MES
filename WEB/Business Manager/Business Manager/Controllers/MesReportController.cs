using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using MESDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;
using Xceed.Words.NET;
using System.Dynamic;
using Newtonsoft.Json.Converters;

namespace Business_Manager.Controllers
{
    public class MesReportController : WebController
    {
        private MesReportDA mesReportDA = new MesReportDA();

        #region //View
        public ActionResult ProductionInProcess()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductionHistory()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BarcodeTracking()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MeasurementRecord()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameProductionHistory()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameBarcodeTracking()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameMeasurementRecord()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameQtyProcess()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductionComponent()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ToolTransactions()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult DeptA372()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MirrorCoatingMoldInformation()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BarcodeMergeHistory()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductQuery()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult IFrameProductQuery()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UserEventHistoryQuery()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MachineEventHistoryQuery()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProcessEventHistoryQuery()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult IFrameBarcodeProcessEvent()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BarcodeProcessEvent()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult LetteringChangeData()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ScrapReport()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ExcessReport()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SaleRankingManagement()
        {
            //ViewLoginCheck();

            return View();
        }

        public ActionResult SaleRankingManagementCNY()
        {
            //ViewLoginCheck();

            return View();
        }

        public ActionResult EtergeWoDetail()
        {
            //ViewLoginCheck();

            return View();
        }

        public ActionResult WoProcessProduct()
        {
            //ViewLoginCheck();

            return View();
        }

        public ActionResult MachineOperationDashboard()
        {
            //ViewLoginCheck();

            return View();
        }

        public ActionResult RptMoInformation()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult JobProgressV2()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult OspDetail()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MoProgressDetail()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult OrderStockReport()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult IFrameOrderStockReport()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RptMoInformationForSales()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult CheckOrderAndProduction()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MWEDetailReport()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MoMaterialsReport()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult InProcessGoodsDefectRateReport()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProcessMenReport()
        {
            ViewLoginCheck();
            return View();
        }


        public ActionResult RBPurchaseRequisitionProposal()
        {
            //粗胚建議請購报表
            ViewLoginCheck();
            return View();
        }

        public ActionResult RBAnnualProcurementReport()
        {
            //年購粗胚明細
            ViewLoginCheck();
            return View();
        }

        public ActionResult IOTDataReport()
        {
            //IOTDataReport
            ViewLoginCheck();
            return View();
        }

        public ActionResult DepProfitLossStatement()
        {
            //成本趋势分析报表
            ViewLoginCheck();
            return View();
        }

        public ActionResult CostAllocationTable()
        {
            //成本核算分配表
            ViewLoginCheck();
            return View();
        }

        public ActionResult SalesCostDetails()
        {
            //晶彩銷售成本明細
            ViewLoginCheck();
            return View();
        }

        public ActionResult ShipmentCostDetails()
        {
            //晶彩出貨成本明細
            ViewLoginCheck();
            return View();
        }

        public ActionResult ReimbursementCostDetails()
        {
            //晶彩歸還成本明細
            ViewLoginCheck();
            return View();
        }

        public ActionResult SalesReturnsAllowancesCostDetails()
        {
            //晶彩銷退/折讓成本明細
            ViewLoginCheck();
            return View();
        }

        public ActionResult CustomerProductInquiry()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult OspDetailSummary()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult PurchasePriceTrend()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult BarcodeBinding()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProcessTransferDetail()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ETGMachineDashboard()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MonthlySalesDetailReport()
        {
            //每月業務撈取ERP銷貨明細 
            ViewLoginCheck();
            return View();
        }

        public ActionResult LCMObsoleteStockProvisionDetail()
        {
            //晶彩銷售成本明細
            ViewLoginCheck();
            return View();
        }

        public ActionResult ExcelExportManagement()
        {
            //晶彩銷售成本明細
            ViewLoginCheck();
            return View();
        }

        #endregion

        #region //Get        
        #region //GetMonthlySalesDetailReport 每月業務撈取ERP銷貨明細
        [HttpPost]
        public void GetMonthlySalesDetailReport(string YearMonth = "", string SalesPersonCode = "", string CustomerCode = "",
            string StartDate = "", string EndDate = "", string SalesOrderNo = "", string ProductNo = "",
            string ProductCategory = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("SalesReport", "read");
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMonthlySalesDetailReport(YearMonth, SalesPersonCode, CustomerCode,
                    StartDate, EndDate, SalesOrderNo, ProductNo, ProductCategory, OrderBy, PageIndex, PageSize);
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

        #region //ExcelMonthlySalesDetailReport 每月業務撈取ERP銷貨明細匯出Excel
        [HttpPost]
        public void ExcelMonthlySalesDetailReport(string YearMonth = "", string SalesPersonCode = "", string CustomerCode = "",
            string StartDate = "", string EndDate = "", string SalesOrderNo = "", string ProductNo = "",
            string ProductCategory = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("SalesReport", "export");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMonthlySalesDetailReport(YearMonth, SalesPersonCode, CustomerCode,
                    StartDate, EndDate, SalesOrderNo, ProductNo, ProductCategory, OrderBy, PageIndex, PageSize);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】每月業務撈取ERP銷貨明細";
                    string excelsheetName = "每月業務銷貨明細";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "單據類別", "銷售日期", "產品分類", "客戶", "業務人員", "銷貨單號", "品號", "品名", "銷貨數量", "本幣金額", "台幣金額" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.單據類別.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.銷售日期.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Style = dateStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.產品分類.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.客戶.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.業務人員.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.銷貨單號.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.品號.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.品名.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.銷貨數量.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Style = numberStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.本幣金額.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Style = currencyStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.台幣金額.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Style = currencyStyle;
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 12;  // 單據類別
                        worksheet.Column(2).Width = 16;  // 銷售日期
                        worksheet.Column(3).Width = 14;  // 產品分類
                        worksheet.Column(4).Width = 20;  // 客戶
                        worksheet.Column(5).Width = 14;  // 業務人員
                        worksheet.Column(6).Width = 16;  // 銷貨單號
                        worksheet.Column(7).Width = 16;  // 品號
                        worksheet.Column(8).Width = 25;  // 品名
                        worksheet.Column(9).Width = 12;  // 銷貨數量
                        worksheet.Column(10).Width = 14; // 本幣金額
                        worksheet.Column(11).Width = 14; // 台幣金額
                        #endregion

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

        #region //GetSalesReportBasicData 取得銷售報表基礎資料(業務人員、客戶、產品分類等下拉選單)
        [HttpPost]
        public void GetSalesReportBasicData(string DataType = "", string UserNo = "")
        {
            try
            {
                //WebApiLoginCheck("SalesReport", "read");
                #region //Request
                mesReportDA = new MesReportDA();

                switch (DataType.ToUpper())
                {
                    case "SALESPERSON":
                        dataRequest = mesReportDA.GetSalesPersonList(UserNo);
                        break;
                    case "CUSTOMER":
                        dataRequest = mesReportDA.GetCustomerList(UserNo);
                        break;
                    case "PRODUCTCATEGORY":
                        dataRequest = mesReportDA.GetProductCategoryList(UserNo);
                        break;
                    default:                        
                        break;
                }
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

        #region //GetProductQuery 取得生產在製資料
        [HttpPost]
        public void GetProductQuery(int MoId = -1, int ModeId = -1, string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", string ProcessAlias = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductQuery(MoId, ModeId, WoErpFullNo, MtlItemNo, StartDate, EndDate, ProcessAlias
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

        #region //GetProductionInProcess 取得生產在製資料
        [HttpPost]
        public void GetProductionInProcess(int ModeId = -1, string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", string CompanyNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductionInProcess", "read");

                #region //切換公司別
                if (CompanyNo.Length > 0)
                {
                    UpdateCompanySwitchForApi(CompanyNo);
                }
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionInProcess(ModeId, WoErpFullNo, MtlItemNo, StartDate, EndDate
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

        #region //GetBarcodeTracking 取得生產在製資料
        [HttpPost]
        public void GetBarcodeTracking(int BarcodeId = -1, string BarcodeNo = "", string MoProcessList = "", string QueryType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("BarcodeTracking", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeTracking(BarcodeId, BarcodeNo, MoProcessList, QueryType
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetProductionHistory 取得生產在製資料
        [HttpPost]
        public void GetProductionHistory(int MoProcessId = -1, string ProdStatus = "")
        {
            try
            {
                WebApiLoginCheck("ProductionHistory", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionHistory(MoProcessId, ProdStatus);
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

        #region //GetWipBarcodeCountData 取得條碼在製資料筆數
        [HttpPost]
        public void GetWipBarcodeCountData(string MoProcessIdListString = "")
        {
            try
            {
                WebApiLoginCheck("ProductionHistory", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWipBarcodeCountData(MoProcessIdListString);
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

        #region //GetQcRecord 取得量測資料
        [HttpPost]
        public void GetQcRecord(int QcRecordId = -1, string DocErpFullNo = "", string QcNoticeNo = "", string MtlItemNo = "", string MtlItemName = ""
            , int QcTypeId = -1, string CheckQcMeasureData = "", int UserId = -1, string StartDate = "", string EndDate = "", string BarcodeNo = "", string QcType = "", int MoProcessId = -1, string LotNumber = "", string QcLotNumber = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcRecord(QcRecordId, DocErpFullNo, QcNoticeNo, MtlItemNo, MtlItemName
                    , QcTypeId, CheckQcMeasureData, UserId, StartDate, EndDate, BarcodeNo, QcType, MoProcessId, LotNumber, QcLotNumber
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
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

        #region //GetQcMeasureData 取得量測詳細資料
        [HttpPost]
        public void GetQcMeasureData(int QcRecordId = -1, string BarcodeNo = "", string QcResult = "", string QcItemName = "", string MachineNumber = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcMeasureData(QcRecordId, BarcodeNo, QcResult, QcItemName, MachineNumber);
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

        #region //GetQcRecordFileForPicture 取得量測記錄檔案(限定PNG格式)
        [HttpPost]
        public void GetQcRecordFileForPicture(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcRecordFileForPicture(QcRecordId);
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

        #region //GetQcRecordFile 取得量測記錄檔案
        [HttpPost]
        public void GetQcRecordFile(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcRecordFile(QcRecordId);
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

        #region //GetMeasurementRecord 取得量測記錄資料
        [HttpPost]
        public void GetMeasurementRecord(string BarcodeNo = "", string WoErpFullNo = "", string WoSeq = "", int ProcessId = -1, string QcType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMeasurementRecord(BarcodeNo, WoErpFullNo, WoSeq, ProcessId, QcType
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

        #region //GetTimeout 取得頁面刷新時間
        [HttpPost]
        public void GetTimeout(string TypeSchema)
        {
            try
            {
                WebApiLoginCheck("ProductionInProcess", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetTimeout(TypeSchema);
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

        #region //GetQryProductionHistory 取得數量生產在製資料
        [HttpPost]
        public void GetQryProductionHistory(int MoId = -1, int ProcessId = -1, string ProdStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                // WebApiLoginCheck("GetQryProductionHistory", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQryProductionHistory(MoId, ProcessId, ProdStatus
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

        #region //GetProductionComponent 取得上料生產資料
        [HttpPost]
        public void GetProductionComponent(int MoId = -1, string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = ""
            , string BarcodeNo = "", string BarcodeProcessNo = "", string BarcodeProcessName = "", string BarcodeComponentNo = "", string BarcodeComponentName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductionComponent", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionComponent(MoId, WoErpFullNo, MtlItemNo, MtlItemName
                    , BarcodeNo, BarcodeProcessNo, BarcodeProcessName, BarcodeComponentNo, BarcodeComponentName
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

        #region //GetToolTransaction 取得刀具異動資料
        [HttpPost]
        public void GetToolTransaction(int ToolTransactionsId = -1, string ToolNo = "", string ToolModelNo = ""
            , int ToolClassId = -1, int ToolCategoryId = -1, string TransactionType = ""
            , string StartDate = "", string EndDate = "", string TimeStatus = ""
            , int ToolInventoryId = -1, int ToolLocatorId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolTransactions", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetToolTransaction(ToolTransactionsId, ToolNo, ToolModelNo
                    , ToolClassId, ToolCategoryId, TransactionType
                    , ToolInventoryId, ToolLocatorId
                    , StartDate, EndDate, TimeStatus
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

        #region //GetDeptA372   copy from GetToolTransaction 取得光薄報表資料
        [HttpPost]
        //public void GetDeptA372(int ToolTransactionsId = -1, string ToolNo = "", string ToolModelNo = ""
        //    , int ToolClassId = -1, int ToolCategoryId = -1, string TransactionType = ""
        //    , string StartDate = "", string EndDate = "", string TimeStatus = ""
        //    , int ToolInventoryId = -1, int ToolLocatorId = -1
        //    , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        public void GetDeptA372(string EndDate = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeptA372", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetDeptA372(EndDate, PageIndex, PageSize);
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

        #region//GetMirrorCoatingMold 取得鍍膜塑膠鏡片
        [HttpPost]
        public void GetMirrorCoatingMold(string MtlItemNo = "", string MtlItemName = "", string ErpNo = "", string BarcodeNo = "",
            string MirrorMoldingStartDate = "", string MirrorMoldingEndDate = "",
            string PlatingMoldStartDate = "", string PlatingMoldEndDate = "", int MachineId = -1,
            string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MirrorCoatingMoldInformation", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMirrorCoatingMold(
                    MtlItemNo, MtlItemName, ErpNo, BarcodeNo,
                    MirrorMoldingStartDate, MirrorMoldingEndDate,
                    PlatingMoldStartDate, PlatingMoldEndDate, MachineId,
                    OrderBy, PageIndex, PageSize);
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

        #region //GetBarcodeTransfer 取得條碼轉移紀錄
        [HttpPost]
        public void GetBarcodeTransfer(string FromBarcodeNo = "", string ToBarcodeNo = "", string WoErpFullNo = "", string TransferStartDate = "", string TransferEndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("BarcodeMergeHistory", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeTransfer(FromBarcodeNo, ToBarcodeNo, WoErpFullNo, TransferStartDate, TransferEndDate
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

        #region //GetQcBarcode 取得量測判定資料
        [HttpPost]
        public void GetQcBarcode(int QcRecordId = -1, string BarcodeNo = "", string QcStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcBarcode(QcRecordId, BarcodeNo, QcStatus
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

        #region //GetWipBarcode 取得特定站別在製條碼
        [HttpPost]
        public void GetWipBarcode(int MoProcessId, string WipStatus
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWipBarcode(MoProcessId, WipStatus
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

        #region //GetUserEventHistory 取得人員事件歷程
        [HttpPost]
        public void GetUserEventHistory(string DepartmentName = "", string UserName = "", string StartDate = "", string FinishDate = "", string LastModifiedDate = "", int LastModifiedBy = -1, string LastModifiedUserName = "", int DepartmentId = -1, int UserId = -1, string UserEventItemName = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserEventHistoryQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetUserEventHistory(DepartmentName, UserName, StartDate, FinishDate, LastModifiedDate, LastModifiedBy, LastModifiedUserName, DepartmentId, UserId, UserEventItemName, OrderBy, PageIndex, PageSize);
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

        #region //GetMachineEventHistory 取得機台事件歷程
        [HttpPost]
        public void GetMachineEventHistory(string MachineName = "", string MachineDesc = "", string ShopName = "", string DuringTime = "", string UserName = "", string CreateDate = "", string StartDate = "", string FinishDate = "", int ShopId = -1, int MachineId = -1, string MachineEventName = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineEventHistoryQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMachineEventHistory(MachineName, MachineDesc, ShopName, DuringTime, UserName, CreateDate, StartDate, FinishDate, ShopId, MachineId, MachineEventName, OrderBy, PageIndex, PageSize);
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

        #region //GetProcessEventHistory 取得加工事件歷程
        [HttpPost]
        public void GetProcessEventHistory(string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string FinishDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ProcessEventHistoryQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProcessEventHistory(WoErpFullNo, MtlItemNo, StartDate, FinishDate, OrderBy, PageIndex, PageSize);
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

        #region //GetBarcodeProcessEvent 取得條碼過站事件資料
        [HttpPost]
        public void GetBarcodeProcessEvent(int UserId = -1, string BarcodeNo = "", string ProcessEventName = "", string MtlItemNo = "", string StartDate = "", string FinishDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ProcessEventHistoryQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeProcessEvent(UserId, BarcodeNo, ProcessEventName, MtlItemNo, StartDate, FinishDate, OrderBy, PageIndex, PageSize);
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

        #region //GetLetteringChangeData 取得NG模仁刻號置換紀錄
        [HttpPost]
        public void GetLetteringChangeData(int ProcessAttrLogId = -1, string WoErpFullNo = "", string MtlItemName = ""
            , string BarcodeNo = "", string OriItemValue = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LetteringChangeData", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetLetteringChangeData(ProcessAttrLogId, WoErpFullNo, MtlItemName
                    , BarcodeNo, OriItemValue, Status
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

        #region //GetScrapReport 品號報廢報表
        [HttpPost]
        public void GetScrapReport(string MtlItemNo = "", string startDate = "", string endDate = "", string WoErpPrefix = "", string WoErpNo = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ScrapReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetScrapReport(MtlItemNo, startDate, endDate, WoErpPrefix, WoErpNo, OrderBy, PageIndex, PageSize);
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

        #region //GetExcessReport 品號報廢報表
        [HttpPost]
        public void GetExcessReport(string MtlItemNo = "", string startDate = "", string endDate = "", string WoErpPrefix = "", string WoErpNo = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ScrapReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetExcessReport(MtlItemNo, startDate, endDate, WoErpPrefix, WoErpNo, OrderBy, PageIndex, PageSize);
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

        #region//GetSaleRankingListVe1 --取得業務成績排行(BY單據業務)
        [HttpPost]
        public void GetSaleRankingListVe1(string Type = "", string Model = "", string StartDate = "", string EndDate = ""
            , string CompanyList = "", int PgId = -1, string ExcludeStatus = "")
        {
            try
            {
                string ClientIP = BaseHelper.ClientIP();
                string AllowStatus = "N";
                #region //Request 
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetAllowUserIp(ClientIP);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    AllowStatus = jsonResponse["result"][0].ToString();
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                if (AllowStatus == "Y")
                {

                }
                else
                {
                    WebApiLoginCheck("SaleRankingManagement", "read");
                }
                #region //Request 
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSaleRankingListVe1(Type, Model, StartDate, EndDate
                    , CompanyList, PgId, ExcludeStatus);
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

        #region//GetSaleRankingListVe2 --取得業務成績排行(BY客戶業務)
        [HttpPost]
        public void GetSaleRankingListVe2(string Type = "", string Model = "", string StartDate = "", string EndDate = ""
            , string CompanyList = "", int PgId = -1, string ExcludeStatus = "")
        {
            try
            {
                string ClientIP = BaseHelper.ClientIP();
                string AllowStatus = "N";
                #region //Request 
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetAllowUserIp(ClientIP);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    AllowStatus = jsonResponse["result"][0].ToString();
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                if (AllowStatus == "Y")
                {

                }
                else
                {
                    WebApiLoginCheck("SaleRankingManagement", "read");
                }
                #region //Request 
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSaleRankingListVe2(Type, Model, StartDate, EndDate
                    , CompanyList, PgId, ExcludeStatus);
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

        #region //GetWoData --基礎過站資訊
        [HttpPost]
        public void GetWoData(int ProcessId = -1, string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", int UserId = -1, string BarcodeNo = "", string Status = "", string isNewFinish = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ScrapReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWoData(ProcessId, WoErpFullNo, MtlItemNo, StartDate, EndDate, UserId, BarcodeNo, Status, isNewFinish
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

        #region //GetProductManufactureData --日夜班產出統計(良品 & 不良品&良率)
        [HttpPost]
        public void GetProductManufactureData(int ModeId = -1, string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ScrapReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductManufactureData(ModeId, WoErpFullNo, MtlItemNo, StartDate, EndDate
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

        #region //GetInspectionInfo 取得送測彙整資訊
        [HttpPost]
        public void GetInspectionInfo(string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetInspectionInfo(StartDate, EndDate);
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

        #region //GetMoProcess 取得製令製程資料
        [HttpPost]
        public void GetMoProcess(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMoProcess(MoId);
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


        #region //GetMachineOperatingTime 取得機器運作時間(分鐘)
        [HttpPost]
        public void GetMachineOperatingTime(int ShopId = -1, int MachineId = -1, string StartDate = "")
        {
            try
            {
                WebApiLoginCheck("MachineOperationDashboard", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMachineOperatingTime(ShopId, MachineId, StartDate);
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

        #region //GetWorkShop 取得車間資料
        [HttpPost]
        public void GetWorkShop(int ShopId = -1, string ShopNo = "", string ShopName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("MachineOperationDashboard", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWorkShop(ShopId, ShopNo, ShopName, Status);
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

        #region //GetWorkShopMachine 取得特定車間內的機台
        [HttpPost]
        public void GetWorkShopMachine(int MachineId = -1, int ShopId = -1, string MachineNo = "", string MachineName = "", string MachineDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("MachineOperationDashboard", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWorkShopMachine(MachineId, ShopId, MachineNo, MachineName, MachineDesc, Status);
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

        #region //GetMachineStatus 取得特定機台稼動狀況
        [HttpPost]
        public void GetMachineStatus(int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("MachineOperationDashboard", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMachineStatus(MachineId);
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

        #region //GetMachineEffectivenessList 取得特定車間機台稼動狀況
        [HttpPost]
        public void GetMachineEffectivenessList(int ShopId = -1, string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("MachineOperationDashboard", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMachineEffectivenessList(ShopId, StartDate, EndDate);
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

        #region //GetRptMoInformation 取得生產異常資料
        [HttpPost]
        public void GetRptMoInformation(string WoErpFullNo = "", int ModeId =-1, string CustomerShortName = "", string SearchKey = "", string ExpectedStart = "", string ExpectedEnd = "", string CompanyNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //切換公司別
                if (CompanyNo.Length > 0)
                {
                    UpdateCompanySwitchForApi(CompanyNo);
                }
                #endregion

                WebApiLoginCheck("RptMoInformation", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRptMoInformation(WoErpFullNo, ModeId, CustomerShortName, SearchKey, ExpectedStart, ExpectedEnd
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetComponentTracking 取得生產異常資料
        [HttpPost]
        public void GetComponentTracking(int BarcodeId = -1)
        {
            try
            {

                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetComponentTracking(BarcodeId);
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

        #region //GetMoUnitSearch --製令串查
        [HttpPost]
        public void GetMoUnitSearch(string MoFullNo = "", int WoId = -1)
        {
            try
            {

                //WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMoUnitSearch(MoFullNo, WoId);
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

        #region //GetOspDetailSummary --托外單交期追蹤
        [HttpPost]
        public void GetOspDetailSummary(string OspNo = "", string SupplierNo = "", string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = ""
            , string OspCreatStartDate = "", string OspCreatEndDate = "", string Status = "", string DelayStatus = "", int UserId = -1, int DepartmentId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                WebApiLoginCheck("OspDetailSummary", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetOspDetailSummary(OspNo, SupplierNo, WoErpFullNo, MtlItemNo, MtlItemName
                    , OspCreatStartDate, OspCreatEndDate, Status, DelayStatus, UserId, DepartmentId
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

        #region //GetOutsourcingDetail --托外單明細
        [HttpPost]
        public void GetOutsourcingDetail(string OspStartDate = "", string OspEndDate = "", string OspCreatStartDate = "", string OspCreatEndDate = "", string OspNo = "", string SupplierNo = "", string SupplierShortName = "", string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = "", int ProcessAlias = -1, string BackSttus = "", string OsrErpFullNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                //WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetOutsourcingDetail(OspStartDate, OspEndDate, OspCreatStartDate, OspCreatEndDate, OspNo, SupplierNo, SupplierShortName, WoErpFullNo, MtlItemNo, MtlItemName, ProcessAlias, BackSttus, OsrErpFullNo, OrderBy, PageIndex, PageSize);
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

        #region //GeMoProgressDetail --製令加工進度
        [HttpPost]
        public void GeMoProgressDetail(string StartDate = "", string EndDate = "", string WoErpFullNo = "", string WoStatus = "", int ProdMode = -1
            , string MtlItemNo = "", string MtlItemName = "", string MoStatus = "", string QuantityStatusQ = "", string ReceiptStatusQ = "", int MoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                //WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GeMoProgressDetail(StartDate, EndDate, WoErpFullNo, WoStatus, ProdMode, MtlItemNo, MtlItemName
                    , MoStatus, QuantityStatusQ, ReceiptStatusQ, MoId
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

        #region //GetOrderStockReport --訂單進度&庫存數量統計報表
        [HttpPost]
        public void GetOrderStockReport(string LensModel = "", string LensModelName = "",string Unshipped = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OrderStockReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetOrderStockReport(LensModel, LensModelName, Unshipped
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

        #region //GetItemOrderReport --品號訂單數量統計報表
        [HttpPost]
        public void GetItemOrderReport(string MtlItemNo = "",string SaleNo=""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OrderStockReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetItemOrderReport(MtlItemNo, SaleNo
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

        #region //GetRptMoInformation2 -- 取得生產異常資料(業務用)
        [HttpPost]
        public void GetRptMoInformation2(string WoErpFullNo = "", string ModeDesc = "", string CustomerShortName = "", string SearchKey = "", string ExpectedStart = "", string ExpectedEnd = "", string CompanyNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //切換公司別
                if (CompanyNo.Length > 0)
                {
                    UpdateCompanySwitchForApi(CompanyNo);
                }
                #endregion

                WebApiLoginCheck("RptMoInformationForSales", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRptMoInformation2(WoErpFullNo, ModeDesc, CustomerShortName, SearchKey, ExpectedStart, ExpectedEnd
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                jsonResponse = JObject.Parse(dataRequest);
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

        #region /GetCheckOrderAndProduction 取得訂單與在製狀況
        [HttpPost]
        public void GetCheckOrderAndProduction(string MtlItemNo = "")
        {
            try
            {
                //WebApiLoginCheck("RptMoInformation", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetCheckOrderAndProduction(MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                //jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetMWEDetailReport 取得庫存明細
        [HttpPost]
        public void GetMWEDetailReport(string TypeOne = "", string TypeTwo = "", string MtlItemNo = "", string StartDate = "", string FinishDate = "")
        {
            try
            {
                //WebApiLoginCheck("RptMoInformation", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMWEDetailReport("", TypeOne, TypeTwo, MtlItemNo, StartDate, FinishDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                //jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetMoMaterialsReport 取得製令材料再製明細
        [HttpPost]
        public void GetMoMaterialsReport(string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = "", string WoErpPrefix = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("RptMoInformation", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMoMaterialsReport("",MtlItemNo, MtlItemName, WoErpFullNo, WoErpPrefix, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                //jsonResponse = JObject.Parse(dataRequest);
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

        #region//GetMWEDetailReportExcel 取得庫存明細表

        public void GetMWEDetailReportExcel(string TypeOne, string TypeTwo, string MtlItemNo, string StartDate, string FinishDate)
        {
            try
            {
                //WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMWEDetailReport("", TypeOne, TypeTwo, MtlItemNo, StartDate, FinishDate);

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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "庫存明細";
                    string excelsheetName = "庫存明細";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "品號", "品名", "規格", "單位", "庫別", "庫別名稱", "庫存數量"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = excelsheetName;
                        var titleRange = worksheet.Range("A1:G1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        #region //圖片

                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;

                        foreach (var item in data)
                        {
                            rowIndex++;

                            startIndex = rowIndex;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemSpec.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.UmoNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MWENo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MWEName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.InventoryQuantity.ToString().Trim();
                            //rowIndex++;

                        }



                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 36;
                        //worksheet.Column(2).Width = 36;
                        #endregion

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

        #region//SendMWEDetailReportExcel 寄送庫存明細表
        [HttpPost]
        [Route("api/MES/SendMWEDetailReportExcel")]
        public void SendMWEDetailReportExcel(string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "SendMWEDetailReportExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMWEDetailReport(Company, "", "", "", EndReceiptDate, StartReceiptDate);

                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {

                    byte[] excelBytes;
                    string fileName = $"庫存明細表-{DateTime.Now:yyyyMMdd}.xlsx";

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
                    defaultStyle.Font.FontName = "微軟正黑體";
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
                    titleStyle.Font.FontName = "微軟正黑體";
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
                    headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                    headerStyle.Font.FontName = "微軟正黑體";
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dataStyle
                    var dataStyle = XLWorkbook.DefaultStyle;
                    dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.TopBorderColor = XLColor.Black;
                    dataStyle.Border.BottomBorderColor = XLColor.Black;
                    dataStyle.Border.LeftBorderColor = XLColor.Black;
                    dataStyle.Border.RightBorderColor = XLColor.Black;
                    dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dataStyle.Font.FontName = "微軟正黑體";
                    dataStyle.Font.FontSize = 12;
                    dataStyle.Font.Bold = false;
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
                    numberStyle.Font.FontName = "微軟正黑體";
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0.00";
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
                    currencyStyle.Font.FontName = "微軟正黑體";
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0.00";
                    #endregion
                    #endregion

                    #region //EXCEL

                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet = workbook.Worksheets.Add("庫存明細表");
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = "庫存明細表";
                        var titleRange = worksheet.Range("A1:G1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        // 設定表頭
                        string[] headers = new string[] { "品號", "品名", "規格", "單位", "庫別", "庫別名稱", "庫存數量"};

                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style = headerStyle;
                        }

                        // 寫入數據
                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                        int rowIndex = 2;

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemSpec.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.UmoNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MWENo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MWEName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.InventoryQuantity.ToString().Trim();
                            //rowIndex++;

                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion



                        #endregion

                        #region 輸出檔案
                        using (MemoryStream ms = new MemoryStream())
                        {
                            workbook.SaveAs(ms);
                            excelBytes = ms.ToArray();
                        }
                        #endregion
                    }
                    #endregion

                    #region//Mail 寄送托外進貨週期性報表

                    //取得Mail樣板設定
                    string SettingSchema = "MWEDetailReportExcel";
                    string SettingNo = "Y";

                    mesReportDA = new MesReportDA();
                    string mailSettingsJson = mesReportDA.SendQcMail(SettingSchema, SettingNo);
                    if (string.IsNullOrEmpty(mailSettingsJson))
                    {
                        throw new Exception("無法取得郵件設定");
                    }

                    var parsedSettings = JObject.Parse(mailSettingsJson);
                    if (parsedSettings["status"].ToString() != "success")
                    {
                        throw new Exception(parsedSettings["msg"].ToString());
                    }

                    dynamic mailSettings = parsedSettings["data"];

                    foreach (var item in mailSettings)
                    {
                        var mailFile = new MailFile
                        {
                            FileName = Path.GetFileNameWithoutExtension(fileName),
                            FileExtension = ".xlsx",
                            FileContent = excelBytes
                        };


                        #region //寄送Mail
                        var mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = $"庫存明細表_{DateTime.Now:yyyy/MM/dd}",
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = item.MailTo,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = $"附件為中揚庫存明細週期性報表，日期:{DateTime.Now:yyyy/MM/dd}",
                            TextBody = "-",
                            FileInfo = new List<MailFile> { mailFile },
                            QcFileFlag = "N"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion

                    Response.Write(JObject.FromObject(new
                    {
                        status = "success",
                        msg = "報表已成功生成並寄送"
                    }).ToString());
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

        #region//GetMoMaterialsReportExcel 取得製令材料再製明細

        public void GetMoMaterialsReportExcel(string MtlItemNo, string MtlItemName, string WoErpFullNo, string WoErpPrefix, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMoMaterialsReport("",MtlItemNo, MtlItemName, WoErpFullNo, WoErpPrefix, "", -1, -1);

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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "製令在製材料明細表";
                    string excelsheetName = "製令在製材料明細表";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "製令編號", "產品品號", "品名", "規格", "開工日", "急料", "加工廠商", "廠商名稱", "生產線別", "線別名稱"
                        , "預計產量", "已生產量", "預計產包裝數", "已生產包裝數", "訂單單號", "母製令號", "材料品號", "單位", "品名", "規格", "製程代碼", "被取替代品號", "被取替代製程"
                        , "在製數量", "在製標準成本", "單位標準成本", "已發料量", "已耗用量", "在製包裝數量", "包裝單位", "已耗用包裝量", "生產庫別", "庫別名稱" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = excelsheetName;
                        var titleRange = worksheet.Range("A1:K1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        #region //圖片

                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;

                        foreach (var item in data)
                        {
                            rowIndex++;

                            startIndex = rowIndex;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.TA001.ToString() + "-" + item.TA002.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.TA006.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.TA034.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.TA035.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.TA009.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.TA044.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.TA032.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.TA032.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.TA021.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.TA021.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.TA015.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.TA017.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.TA045.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.TA046.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.TA026C.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.TA024C.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.TB003.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.TB007.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.TB012.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.TB013.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.TB006.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.TB023.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.TB025.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.InProcessQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.MB060C.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.TA022.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.TB005.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.usedQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.TA046.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.TA048.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.TA046.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.TA020.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.MC002.ToString();

                            //rowIndex++;

                        }



                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;
                        #endregion

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

        #region//SendMoMaterialsReportExcel 寄送 製令材料再製明細
        [HttpPost]
        [Route("api/MES/SendMoMaterialsReportExcel")]
        public void SendMoMaterialsReportExcel(string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "SendMoMaterialsReportExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyyMMdd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyyMMdd");

                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMoMaterialsReport(Company, "", "", "", "", "", -1, -1);

                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {

                    byte[] excelBytes;
                    string fileName = $"製令材料在製明細-{DateTime.Now:yyyyMMdd}.xlsx";

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
                    defaultStyle.Font.FontName = "微軟正黑體";
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
                    titleStyle.Font.FontName = "微軟正黑體";
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
                    headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                    headerStyle.Font.FontName = "微軟正黑體";
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dataStyle
                    var dataStyle = XLWorkbook.DefaultStyle;
                    dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.TopBorderColor = XLColor.Black;
                    dataStyle.Border.BottomBorderColor = XLColor.Black;
                    dataStyle.Border.LeftBorderColor = XLColor.Black;
                    dataStyle.Border.RightBorderColor = XLColor.Black;
                    dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dataStyle.Font.FontName = "微軟正黑體";
                    dataStyle.Font.FontSize = 12;
                    dataStyle.Font.Bold = false;
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
                    numberStyle.Font.FontName = "微軟正黑體";
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0.00";
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
                    currencyStyle.Font.FontName = "微軟正黑體";
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0.00";
                    #endregion
                    #endregion

                    #region //EXCEL

                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet = workbook.Worksheets.Add("製令材料在製明細");
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = "製令材料在製明細";
                        var titleRange = worksheet.Range("A1:K1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        // 設定表頭
                        string[] headers = new string[] { "製令編號", "產品品號", "品名", "規格", "開工日", "急料", "加工廠商", "廠商名稱", "生產線別", "線別名稱"
                        , "預計產量", "已生產量", "預計產包裝數", "已生產包裝數", "訂單單號", "母製令號", "材料品號", "單位", "品名", "規格", "製程代碼", "被取替代品號", "被取替代製程"
                        , "在製數量", "在製標準成本", "單位標準成本", "已發料量", "已耗用量", "在製包裝數量", "包裝單位", "已耗用包裝量", "生產庫別", "庫別名稱" };


                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style = headerStyle;
                        }

                        // 寫入數據
                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                        int rowIndex = 2;

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.TA001.ToString() + "-" + item.TA002.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.TA006.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.TA034.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.TA035.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.TA009.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.TA044.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.TA032.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.TA032.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.TA021.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.TA021.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.TA015.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.TA017.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.TA045.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.TA046.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.TA026C.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.TA024C.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.TB003.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.TB007.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.TB012.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.TB013.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.TB006.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.TB023.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.TB025.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.InProcessQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.MB060C.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.TA022.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.TB005.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.usedQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.TA046.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.TA048.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.TA046.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.TA020.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.MC002.ToString();
                            //rowIndex++;

                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion



                        #endregion

                        #region 輸出檔案
                        using (MemoryStream ms = new MemoryStream())
                        {
                            workbook.SaveAs(ms);
                            excelBytes = ms.ToArray();
                        }
                        #endregion
                    }
                    #endregion

                    #region//Mail 寄送托外進貨週期性報表

                    //取得Mail樣板設定
                    string SettingSchema = "MoMaterialsReportExcel";
                    string SettingNo = "Y";

                    mesReportDA = new MesReportDA();
                    string mailSettingsJson = mesReportDA.SendQcMail(SettingSchema, SettingNo);
                    if (string.IsNullOrEmpty(mailSettingsJson))
                    {
                        throw new Exception("無法取得郵件設定");
                    }

                    var parsedSettings = JObject.Parse(mailSettingsJson);
                    if (parsedSettings["status"].ToString() != "success")
                    {
                        throw new Exception(parsedSettings["msg"].ToString());
                    }

                    dynamic mailSettings = parsedSettings["data"];

                    foreach (var item in mailSettings)
                    {
                        var mailFile = new MailFile
                        {
                            FileName = Path.GetFileNameWithoutExtension(fileName),
                            FileExtension = ".xlsx",
                            FileContent = excelBytes
                        };


                        #region //寄送Mail
                        var mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = $"製令材料在製明細_{DateTime.Now:yyyy/MM/dd}",
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = item.MailTo,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = $"附件為製令材料在製明細週期性報表，日期:{DateTime.Now:yyyy/MM/dd}",
                            TextBody = "-",
                            FileInfo = new List<MailFile> { mailFile },
                            QcFileFlag = "N"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion

                    Response.Write(JObject.FromObject(new
                    {
                        status = "success",
                        msg = "報表已成功生成並寄送"
                    }).ToString());
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

        #region //GetInProcessGoodsDefectRateReport --在製品不良率報表
        [HttpPost]
        public void GetInProcessGoodsDefectRateReport(int ModeId = -1, string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        {
            try
            {
                WebApiLoginCheck("InProcessGoodsDefectRateReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetInProcessGoodsDefectRateReport(ModeId, MtlItemNo, StartDate, EndDate
                    , OrderBy, PageIndex, PageSize, RedirectType);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetProcessMenReport 取得人員製造狀況報表
        [HttpPost]
        public void GetProcessMenReport(string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = ""
            , int MachineId = -1, int UserId = -1, int ModeId = -1, int ProcessId = -1, string StartDate = "", string FinishDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("RptMoInformation", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProcessMenReport(MtlItemNo, MtlItemName, WoErpFullNo, MachineId, UserId, ModeId, ProcessId, StartDate, FinishDate, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                //jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetRBPurchaseRequisitionProposal 粗胚建議請購报表
        [HttpPost]
        public void GetRBPurchaseRequisitionProposal(string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RBPurchaseRequisitionProposal", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRBPurchaseRequisitionProposal(MtlItemNo, MtlItemName
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

        #region //GetRBAnnualProcurementReport 年購粗胚明細
        [HttpPost]
        public void GetRBAnnualProcurementReport(string StartReqDate = "", string EndReqDate = "", string ReqName = ""
                    , string ReqNo = "", string ReqDepNo = "", string ReqUserNo = "", string SearchKey = ""
                    , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
                    , string PurUserNo = "", string PurVendorName = ""
                    , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RBAnnualProcurementReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRBAnnualProcurementReport(StartReqDate, EndReqDate, ReqName
                    , ReqNo, ReqDepNo, ReqUserNo, SearchKey
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
                    , PurUserNo, PurVendorName
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

        #region //GetIOTData -- 取得自動化資料表數據
        [HttpPost]
        public void GetIOTData(string MachineType = "", string StartDate = "", string FinishDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MesReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, StartDate, FinishDate
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

        #region //GetIOTData2 -- 取得自動化資料表數據
        [HttpPost]
        public void GetIOTData2(string MachineType = "", string StartDate = "", string FinishDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MesReport", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData2(MachineType, StartDate, FinishDate
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


        #region//SendMWEDetailReportExcel 寄送庫存明細表(紘立)
        [HttpPost]
        [Route("api/MES/SendETGMWEDetailReportExcel")]
        public void SendETGMWEDetailReportExcel(string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "SendETGMWEDetailReportExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMWEDetailReport(Company, "", "", "", EndReceiptDate, StartReceiptDate);

                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {

                    byte[] excelBytes;
                    string fileName = $"庫存明細表-{DateTime.Now:yyyyMMdd}.xlsx";

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
                    defaultStyle.Font.FontName = "微軟正黑體";
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
                    titleStyle.Font.FontName = "微軟正黑體";
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
                    headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                    headerStyle.Font.FontName = "微軟正黑體";
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dataStyle
                    var dataStyle = XLWorkbook.DefaultStyle;
                    dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.TopBorderColor = XLColor.Black;
                    dataStyle.Border.BottomBorderColor = XLColor.Black;
                    dataStyle.Border.LeftBorderColor = XLColor.Black;
                    dataStyle.Border.RightBorderColor = XLColor.Black;
                    dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dataStyle.Font.FontName = "微軟正黑體";
                    dataStyle.Font.FontSize = 12;
                    dataStyle.Font.Bold = false;
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
                    numberStyle.Font.FontName = "微軟正黑體";
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0.00";
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
                    currencyStyle.Font.FontName = "微軟正黑體";
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0.00";
                    #endregion
                    #endregion

                    #region //EXCEL

                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet = workbook.Worksheets.Add("庫存明細表");
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = "庫存明細表";
                        var titleRange = worksheet.Range("A1:G1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        // 設定表頭
                        string[] headers = new string[] { "品號", "品名", "規格", "單位", "庫別", "庫別名稱", "庫存數量" };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style = headerStyle;
                        }

                        // 寫入數據
                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                        int rowIndex = 2;

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemSpec.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.UmoNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MWENo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MWEName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.InventoryQuantity.ToString().Trim();
                            //rowIndex++;

                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion



                        #endregion

                        #region 輸出檔案
                        using (MemoryStream ms = new MemoryStream())
                        {
                            workbook.SaveAs(ms);
                            excelBytes = ms.ToArray();
                        }
                        #endregion
                    }
                    #endregion

                    #region//Mail 寄送托外進貨週期性報表

                    //取得Mail樣板設定
                    string SettingSchema = "ETGMWEDetailReportExcel";
                    string SettingNo = "Y";

                    mesReportDA = new MesReportDA();
                    string mailSettingsJson = mesReportDA.SendQcMail(SettingSchema, SettingNo);
                    if (string.IsNullOrEmpty(mailSettingsJson))
                    {
                        throw new Exception("無法取得郵件設定");
                    }

                    var parsedSettings = JObject.Parse(mailSettingsJson);
                    if (parsedSettings["status"].ToString() != "success")
                    {
                        throw new Exception(parsedSettings["msg"].ToString());
                    }

                    dynamic mailSettings = parsedSettings["data"];

                    foreach (var item in mailSettings)
                    {
                        var mailFile = new MailFile
                        {
                            FileName = Path.GetFileNameWithoutExtension(fileName),
                            FileExtension = ".xlsx",
                            FileContent = excelBytes
                        };


                        #region //寄送Mail
                        var mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = $"庫存明細表_{DateTime.Now:yyyy/MM/dd}",
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = item.MailTo,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = $"附件為紘立庫存明細週期性報表，日期:{DateTime.Now:yyyy/MM/dd}",
                            TextBody = "-",
                            FileInfo = new List<MailFile> { mailFile },
                            QcFileFlag = "N"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion

                    Response.Write(JObject.FromObject(new
                    {
                        status = "success",
                        msg = "報表已成功生成並寄送"
                    }).ToString());
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

        #region //GetDepProfitLossStatement 成本趋势分析报表
        [HttpPost]
        public void GetDepProfitLossStatement(string Department = "", string FiscalYear = ""
                    , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DepProfitLossStatement", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetDepProfitLossStatement(Department, FiscalYear
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

        #region //GetCostAllocationTable 成本核算分配表
        [HttpPost]
        public void GetCostAllocationTable(string YearMonth = ""
                    , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CostAllocationTable", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetCostAllocationTable(YearMonth
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

        #region //GetSalesCostDetails 晶彩銷售成本明細
        [HttpPost]
        public void GetSalesCostDetails(string RoErpPrefix = "", string RoErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SalesCostDetails", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSalesCostDetails(RoErpPrefix, RoErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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

        #region //GetShipmentCostDetails 晶彩出貨成本明細
        [HttpPost]
        public void GetShipmentCostDetails(string TsnErpPrefix = "", string TsnErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SalesCostDetails", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetShipmentCostDetails(TsnErpPrefix, TsnErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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

        #region //GetReimbursementCostDetails 晶彩歸還成本明細
        [HttpPost]
        public void GetReimbursementCostDetails(string TsrnErpPrefix = "", string TsrnErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SalesCostDetails", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetReimbursementCostDetails(TsrnErpPrefix, TsrnErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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

        #region //GetSalesReturnsAllowancesCostDetails 晶彩銷退/折讓成本明細
        [HttpPost]
        public void GetSalesReturnsAllowancesCostDetails(string RtErpPrefix = "", string RtErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SalesCostDetails", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSalesReturnsAllowancesCostDetails(RtErpPrefix, RtErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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

        #region //GetCustomerProduct 取得指定客戶的機種資料(可能有多個客戶)
        [HttpPost]
        public void GetCustomerProduct(string CustomerNo = "", string ProductModel = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerProductInquiry", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetCustomerProduct(CustomerNo, ProductModel
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetSoDetailByMtlItem 訂單查詢頁面
        [HttpPost]
        public void GetSoDetailByMtlItem(string CustomerNo = "", string ProductModel = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerProductInquiry", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSoDetailByMtlItem(CustomerNo, ProductModel
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetProductionByMtlItem 生產歷程狀況
        [HttpPost]
        public void GetProductionByMtlItem(string ProductModel = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerProductInquiry", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionByMtlItem(ProductModel
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetProductionStockInByMtlItem 生產入庫資料查詢
        [HttpPost]
        public void GetProductionStockInByMtlItem(string CustomerNo = "", string ProductModel = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerProductInquiry", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionStockInByMtlItem(CustomerNo, ProductModel
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetTempShippingNoteByMtlItem 暫出單資料查詢
        [HttpPost]
        public void GetTempShippingNoteByMtlItem(string CustomerNo = "", string ProductModel = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerProductInquiry", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetTempShippingNoteByMtlItem(CustomerNo, ProductModel
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetShippingOrderByMtlItem 銷貨單資料查詢
        [HttpPost]
        public void GetShippingOrderByMtlItem(string CustomerNo = "", string ProductModel = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerProductInquiry", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetShippingOrderByMtlItem(CustomerNo, ProductModel
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetPurchasePriceTrend --商品採購單價漲跌記錄
        [HttpPost]
        public void GetPurchasePriceTrend(string MtlItemNo = "", string MtlItemName = "", string SupplierNo = ""
            , string PriceDown = "", string PriceUp = "", string PriceKeep = ""
            , string ChooseStartDate = "", string ChooseEndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                WebApiLoginCheck("PurchasePriceTrend", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetPurchasePriceTrend(MtlItemNo, MtlItemName, SupplierNo
                    , PriceDown, PriceUp, PriceKeep
                    , ChooseStartDate, ChooseEndDate
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

        #region //GetPurchasePriceTrendSummary --商品採購單價漲跌記錄(彙總)
        [HttpPost]
        public void GetPurchasePriceTrendSummary(string PoPrefix = "", string SupplierNo = "")
        {
            try
            {

                WebApiLoginCheck("PurchasePriceTrend", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetPurchasePriceTrendSummary(PoPrefix, SupplierNo);
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


        #region //GetBarcodeBindingData --條碼綁定資訊 
        [HttpPost]
        public void GetBarcodeBindingData(string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = "", string BatchBarcodeNo = "", string LensBarcodeNo = ""
            , string ProcessAlias = "", string StartDate = "", string EndDate = "", string UserNo = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                //WebApiLoginCheck("PurchasePriceTrend", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeBindingData(MtlItemNo, MtlItemName, WoErpFullNo, BatchBarcodeNo, LensBarcodeNo
            , ProcessAlias, StartDate, EndDate, UserNo, OrderBy, PageIndex, PageSize);
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

        #region //GetAssembleRealTimeStatus 設備即時狀態 
        [HttpPost]
        public void GetAssembleRealTimeStatus()
        {
            try
            {

                //WebApiLoginCheck("PurchasePriceTrend", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetAssembleRealTimeStatus();
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

        #region //GetTransferDetailData --生產移轉明細資訊 
        [HttpPost]
        public void GetTransferDetailData(int ModeId = -1, string WoErpFullNo = "", string MtlItemName = "", string StartDate = "", string EndDate = "",
            int UserId = -1, string BarcodeNo = "", string ProcessAlias = "", string ShopDesc = "", string MachineDesc = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            {
                try
                {

                    // WebApiLoginCheck("PurchasePriceTrend", "read");

                    #region //Request
                    mesReportDA = new MesReportDA();
                    dataRequest = mesReportDA.GetTransferDetailData(ModeId, WoErpFullNo, MtlItemName, StartDate, EndDate,
                 UserId, BarcodeNo, ProcessAlias, ShopDesc, MachineDesc, OrderBy, PageIndex, PageSize);
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
        }

        #endregion

        #region //GetLCMObsoleteStockProvisionDetail JC實時庫存LCM呆滯提列
        [HttpPost]
        public void GetLCMObsoleteStockProvisionDetail(string MtlItemNo = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        {
            try
            {
                WebApiLoginCheck("LCMObsoleteStockProvisionDetail", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetLCMObsoleteStockProvisionDetail(MtlItemNo
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
                    , OrderBy, PageIndex, PageSize, RedirectType);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetIOTQRData  -- 取得自動化資料表數據-紘立MTF -QR 
        [HttpPost]
        public void GetIOTQRData(string QRNos)
        {
            try
            {

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTQRData(QRNos);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //GetPoStatusReport 取得采购进货状况分析表
        [HttpPost]
        public void GetPoStatusReport(string PoFullNo = "", string MtlItemNo = "", string SupplierNo = "", string CloseCode = "", string PoStatus = "", string DeliveryStatus = ""
            , string StardDay = "", string EndDay = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        {
            try
            {
                WebApiLoginCheck("ExcelExportManagement", "read,excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetPoStatusReport(PoFullNo, MtlItemNo, SupplierNo, CloseCode, PoStatus, DeliveryStatus, StardDay, EndDay, OrderBy, PageIndex, PageSize, RedirectType);
                #endregion

                #region //Response
                //jsonResponse= BaseHelper.DAResponse(dataRequest);
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //Update
        #region //UpdateTimeout 更新頁面刷新時間
        [HttpPost]
        public void UpdateTimeout(string TypeSchema = "", string TypeNo = "")
        {
            try
            {
                WebApiLoginCheck("ProductionInProcess", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.UpdateTimeout(TypeSchema, TypeNo);
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

        #region //UpdateExpectedEnd 更新預計完工日
        [HttpPost]
        public void UpdateExpectedEnd(string WipNo = "", string ExpectedEnd = "")
        {
            try
            {
                WebApiLoginCheck("RptMoInformation", "update-date");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.UpdateExpectedEnd(WipNo, ExpectedEnd);
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

        #region //Download

        #region //Excel

        #region//ExcelToolTransaction 刀具異動報表       
        public void ExcelToolTransaction(int ToolTransactionsId = -1, string ToolNo = "", string ToolModelNo = ""
            , int ToolClassId = -1, int ToolCategoryId = -1, string TransactionType = ""
            , string StartDate = "", string EndDate = "", string TimeStatus = ""
            , int ToolInventoryId = -1, int ToolLocatorId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetToolTransaction(ToolTransactionsId, ToolNo, ToolModelNo
                                                            , ToolClassId, ToolCategoryId, TransactionType
                                                            , ToolInventoryId, ToolLocatorId
                                                            , StartDate, EndDate, TimeStatus
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】工具異動資料查詢";
                    string excelsheetName = "工具異動彙整頁";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "工具編號", "工具型號", "類別", "種類名", "異動類別", "異動日期", "異動原因", "倉庫", "儲位", "新增人員", "加工數" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 10).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ToolNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ToolModelNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.ToolClassName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ToolCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.TransactionTypeName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.CreateDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.TransactionReason.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ToolInventoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ToolLocatorName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.UserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.ProcessingQty.ToString();
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;
                        #endregion

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

        #region//ExcelDeptA372 EXCEL光薄報表 , by MarkChen, 2023-09-11              
        public void ExcelDeptA372(string EndDate = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();

                //dataRequest = mesReportDA.GetDeptA372(EndDate, 1, 1, "Y"
                //                        , OrderBy, PageIndex, PageSize);
                dataRequest = mesReportDA.GetDeptA372(EndDate, PageIndex, PageSize);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】光薄報表";
                    string excelsheetName = "日期";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "站別名稱", "站別別名", "日期", "機種", "人員", "機台", "數量", "加工工時(秒)" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 2)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 8).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int sumQty = 0;
                        int sumSec = 0;
                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ProcessName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ProcessAlias.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.StartDateDayOnly.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.StartUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MachineName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.SumQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.SumSec.ToString();
                            if (int.TryParse(item.SumQty.ToString(), out int sumQtyValue))
                            {
                                sumQty += sumQtyValue;
                            }
                            if (int.TryParse(item.SumSec.ToString(), out int sumSecValue))
                            {
                                sumSec += sumSecValue;
                            }
                        }
                        rowIndex++;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = "合計";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = sumQty;
                        worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = sumSec;
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 20;

                        worksheet.Column(7).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        worksheet.Column(8).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);


                        #endregion

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

        #region //ExcelMirrorCoatingMold 鍍膜塑膠鏡片報表
        public void ExcelMirrorCoatingMold(string MtlItemNo = "", string MtlItemName = "", string ErpNo = "", string BarcodeNo = "",
            string MirrorMoldingStartDate = "", string MirrorMoldingEndDate = "",
            string PlatingMoldStartDate = "", string PlatingMoldEndDate = "", int MachineId = -1,
            string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MirrorCoatingMoldInformation", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMirrorCoatingMold(
                    MtlItemNo, MtlItemName, ErpNo, BarcodeNo,
                    MirrorMoldingStartDate, MirrorMoldingEndDate,
                    PlatingMoldStartDate, PlatingMoldEndDate, MachineId,
                    OrderBy, PageIndex, PageSize);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】鍍膜塑膠鏡片報表";
                    string excelsheetName = "製令彙整頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "產品條碼", "製令", "品號", "穴號", "成型日期", "鍍模機台", "鍋次", "檢查日期", "條碼數量" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;
                            var Cell1 = worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex));
                            var Cell2 = worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex));
                            var Cell3 = worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex));
                            var Cell4 = worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex));
                            var Cell5 = worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex));
                            var Cell6 = worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex));
                            var Cell7 = worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex));
                            var Cell8 = worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex));
                            var Cell9 = worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex));

                            Cell1.Value = item.BarcodeNo.ToString();
                            Cell2.Value = item.WoErpFullNo.ToString();
                            Cell3.Value = item.MtlItemNo.ToString();
                            Cell4.Value = item.CavityItemValue.ToString();
                            Cell5.Value = item.CavityFinishDate.ToString();
                            Cell6.Value = item.MachineDesc.ToString();
                            Cell7.Value = item.PotNumberItemValue.ToString();
                            Cell8.Value = item.PotNumberFinishDate.ToString();
                            Cell9.Value = item.BarcodeQty.ToString();

                            Cell7.Style.NumberFormat.Format = "@";
                        }
                        #endregion

                        #region //設定
                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;
                        #endregion

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

        #region//ExcelUserEventHistory 人員事件歷程報表       
        public void ExcelUserEventHistory(string DepartmentName = "", string UserName = "", string StartDate = "", string FinishDate = "", string LastModifiedDate = "", int LastModifiedBy = -1, string LastModifiedUserName = "", int DepartmentId = -1, int UserId = -1, string UserEventItemName = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserEventHistoryQuery", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetUserEventHistory(DepartmentName, UserName, StartDate, FinishDate, LastModifiedDate, LastModifiedBy, LastModifiedUserName, DepartmentId, UserId, UserEventItemName, OrderBy, PageIndex, PageSize);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】人員事件歷程查詢";
                    string excelsheetName = "人員事件歷程彙整頁";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "部門", "事件建立人員", "事件名稱", "開始時間", "結束時間", "費時", "最後修改人員" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DepartmentName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.UserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.UserEventItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.StartDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.FinishDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.DuringTime.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.LastModifiedUserName.ToString();

                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;
                        #endregion

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

        #region//ExcelMachineEventHistory 機台事件歷程報表       
        public void ExcelMachineEventHistory(string MachineName = "", string MachineDesc = "", string ShopName = "", string DuringTime = "", string UserName = "", string CreateDate = "", string StartDate = "", string FinishDate = "", int ShopId = -1, int MachineId = -1, string MachineEventName = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineEventHistoryQuery", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetMachineEventHistory(MachineName = "", MachineDesc = "", ShopName = "", DuringTime = "", UserName = "", CreateDate = "", StartDate = "", FinishDate = "", ShopId = -1, MachineId = -1, MachineEventName = "", OrderBy = "", PageIndex = -1, PageSize = -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】機台事件歷程查詢";
                    string excelsheetName = "機台事件歷程彙整頁";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "機台名稱", "機台描述", "車間", "事件名稱", "開始時間", "結束時間", "費時", "建立人員", "建立日期" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MachineName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MachineDesc.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.ShopName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MachineEventName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.StartDate.ToString("yyyy-MM-dd HH:mm:ss");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.FinishDate.ToString("yyyy-MM-dd HH:mm:ss");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.DuringTime.ToString("HH:mm:ss");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.UserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.CreateDate.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 20;
                        worksheet.Column(2).Width = 22;
                        worksheet.Column(3).Width = 20;
                        worksheet.Column(4).Width = 22;
                        worksheet.Column(5).Width = 20;
                        worksheet.Column(6).Width = 22;
                        worksheet.Column(7).Width = 20;
                        worksheet.Column(8).Width = 22;
                        worksheet.Column(9).Width = 20;
                        #endregion

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

        #region//ExcelProcessEventHistory 加工事件歷程報表       
        public void ExcelProcessEventHistory(string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string FinishDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ProcessEventHistoryQuery", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProcessEventHistory(WoErpFullNo, MtlItemNo, StartDate, FinishDate, OrderBy, PageIndex, PageSize);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】加工事件歷程查詢";
                    string excelsheetName = "加工事件歷程彙整頁";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "製令", "品號", "品名", "製程", "機台", "條碼", "加工類別", "事件", "開始時間", "結束時間", "人員" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.WoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.StartDate.ToString("yyyy-MM-dd HH:mm:ss");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.FinishDate.ToString("yyyy-MM-dd HH:mm:ss");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.BarcodeNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.MachineName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ProcessName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.TypeName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ProcessEventName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.UserName.ToString();
                        }


                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 20;
                        worksheet.Column(2).Width = 22;
                        worksheet.Column(3).Width = 20;
                        worksheet.Column(4).Width = 22;
                        worksheet.Column(5).Width = 20;
                        worksheet.Column(6).Width = 22;
                        worksheet.Column(7).Width = 20;
                        worksheet.Column(8).Width = 22;
                        worksheet.Column(9).Width = 20;
                        worksheet.Column(10).Width = 20;
                        worksheet.Column(11).Width = 20;
                        #endregion

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

        #region//ExcelScrapReport 報廢報表       
        public void ExcelScrapReport(string MtlItemNo = "", string startDate = "", string endDate = "", string WoErpPrefix = "", string WoErpNo = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ScrapReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetScrapReport(MtlItemNo, startDate, endDate, WoErpPrefix, WoErpNo, OrderBy, PageIndex, PageSize);
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
                    defaultStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體的
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】報廢報表";
                    string excelsheetName = "報廢報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    List<ScrapMo> detailReportDatas = JsonConvert.DeserializeObject<List<ScrapMo>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "製令", "狀態", "產品品號", "產品品名", "產品規格", "預計產量", "領料數", "產品已生產量", "產品報廢數"
                                                   , "材料品號", "材料品名", "材料規格", "需領用數", "已領用數", "材料報廢數", "製令結案報廢數"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 7)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 16).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            //rowIndex++;
                            var s = "";
                            switch (item.MoStatus.ToString())
                            {
                                case "Y":
                                case "y":
                                    s = "已結案";
                                    break;
                                default:
                                    s = "未結案";
                                    break;

                            }
                            foreach (var subitem in item.ScrapMoMtlMerge)
                            {

                                rowIndex++;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.WoErpNo.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = s;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MtlItemspec.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ExpectedCount.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ReceiveCount.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ProducedCount.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ScrapCount.ToString();




                                foreach (var subitem2 in subitem.ScrapMoMtldata)
                                {
                                    //rowIndex++;

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = subitem2.MtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = subitem2.MtlItemName.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = subitem2.MtlItemspec.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = subitem2.ExpectedCount.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = subitem2.ReceiveCount.ToString();

                                }

                                var subitemrowIndex = rowIndex;

                                if (subitem.ScrapMoReplaceMtldata.Count() > 0) {
                                    foreach (var subitem3 in subitem.ScrapMoReplaceMtldata)
                                    {
                                        rowIndex++;

                                        worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = subitem3.MtlItemNo.ToString().Trim();
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = subitem3.MtlItemName.ToString();
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = subitem3.MtlItemspec.ToString().Trim();
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = subitem3.ExpectedCount.ToString();
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = subitem3.ReceiveCount.ToString();

                                    }
                                }


                                worksheet.Cell(BaseHelper.MergeNumberToChar(15, subitemrowIndex)).Value = subitem.MergeScrapCount.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(16, subitemrowIndex)).Value = subitem.MergeMoScrapCount.ToString();
                                worksheet.Range(subitemrowIndex, 15, rowIndex, 15).Merge();
                                worksheet.Range(subitemrowIndex, 16, rowIndex, 16).Merge();






                            }

                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                            worksheet.Range(startIndex, 5, rowIndex, 5).Merge();
                            worksheet.Range(startIndex, 6, rowIndex, 6).Merge();
                            worksheet.Range(startIndex, 7, rowIndex, 7).Merge();
                            worksheet.Range(startIndex, 8, rowIndex, 8).Merge();
                            worksheet.Range(startIndex, 9, rowIndex, 9).Merge();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 30;
                        worksheet.Column(2).Width = 15;
                        worksheet.Column(3).Width = 30;
                        worksheet.Column(4).Width = 30;
                        worksheet.Column(5).Width = 30;
                        worksheet.Column(6).Width = 20;
                        worksheet.Column(7).Width = 20;
                        worksheet.Column(8).Width = 20;
                        worksheet.Column(9).Width = 20;
                        worksheet.Column(10).Width = 30;
                        worksheet.Column(11).Width = 30;
                        worksheet.Column(12).Width = 30;
                        worksheet.Column(13).Width = 20;
                        worksheet.Column(14).Width = 20;
                        worksheet.Column(15).Width = 20;


                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelExcessReport 超領報表       
        public void ExcelExcessReport(string MtlItemNo = "", string startDate = "", string endDate = "", string WoErpPrefix = "", string WoErpNo = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetExcessReport(MtlItemNo, startDate, endDate, WoErpPrefix, WoErpNo, OrderBy, PageIndex, PageSize);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】報領報表";
                    string excelsheetName = "超領報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    List<ScrapMoExcessMtl> detailReportDatas = JsonConvert.DeserializeObject<List<ScrapMoExcessMtl>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "製令", "產品品號", "產品品名", "產品規格", "產品預計產量", "產品已生產量", "材料需領用數"
                                                   , "材料已領用數", "超領數", "單位", "超領率"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.WoErpNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemspec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MoExpectedCount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MoProducedCount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ExpectedCount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ReceiveCount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ExcessCount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Unit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.ExcessPerson.ToString();




                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 30;
                        worksheet.Column(2).Width = 30;
                        worksheet.Column(3).Width = 30;
                        worksheet.Column(4).Width = 30;
                        worksheet.Column(5).Width = 30;
                        worksheet.Column(6).Width = 20;
                        worksheet.Column(7).Width = 20;
                        worksheet.Column(8).Width = 20;
                        worksheet.Column(9).Width = 20;
                        worksheet.Column(10).Width = 15;
                        worksheet.Column(11).Width = 15;



                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelGetWoDataa 基礎過站資訊       
        public void ExcelGetWoDataa(int ProcessId = -1, string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", int UserId = -1, string BarcodeNo = "", string Status = "", string isNewFinish = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWoData(ProcessId, WoErpFullNo, MtlItemNo, StartDate, EndDate, UserId, BarcodeNo, Status, isNewFinish
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】基礎過站資訊";
                    string excelsheetName = "基礎過站資訊";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    List<EtergeWoDataDetail> detailReportDatas = JsonConvert.DeserializeObject<List<EtergeWoDataDetail>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "工作日", "製令", "產品品號", "產品品名", "條碼", "製程", "機台", "開工人員", "上製程完工時間", "開工時間", "完工時間", "完工人員", "過站數量", "條碼狀態", "不良原因", "加工時間/s", "等待時間/s", "是否最新過站" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.WorkDay.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.WoErpFullNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.BarcodeNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ProcessAlias.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.MachineDesc.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.StarUser.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.LastStopFinish != null ? item.LastStopFinish.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.StartDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.FinishDate != null ? item.FinishDate.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.FinishUser != null ? item.FinishUser.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.StationQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.BarcodeStatus.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.CauseNo != null ? item.CauseNo.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.CycleTime.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.WaitTime.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.isNewFinish.ToString();





                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 30;
                        //worksheet.Column(2).Width = 30;
                        //worksheet.Column(3).Width = 30;
                        //worksheet.Column(4).Width = 30;
                        //worksheet.Column(5).Width = 30;
                        //worksheet.Column(6).Width = 20;
                        //worksheet.Column(7).Width = 20;
                        //worksheet.Column(8).Width = 20;
                        //worksheet.Column(9).Width = 20;
                        //worksheet.Column(10).Width = 15;
                        //worksheet.Column(11).Width = 15;



                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //DownloadInspectionExcel 送測彙整EXCEL
        public void DownloadInspectionExcel(string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetInspectionInfo(StartDate, EndDate);
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
                    headerStyle.Fill.BackgroundColor = XLColor.AppleGreen;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
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
                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "送測彙整清單";
                    string excelsheetName = "送測彙整清單詳細資料";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] headerValues = new string[] { "量測單號", "製令", "生產模式", "製程", "送測時間", "確認/駁回時間", "量測完成時間", "量測單據狀況", "駁回原因", "送測人員", "送測部門" };
                    string colIndexValue = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 4)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 8).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < headerValues.Length; i++)
                        {
                            colIndexValue = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndexValue).Value = headerValues[i];
                            worksheet.Cell(colIndexValue).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.QcRecordId.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.WoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.ModeFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ProcessAlias.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.RequestDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ReceiptDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.QcFinishDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.CheckQcMeasureData.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.DisallowanceReason.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.UserFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.DepartmentFullNo.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = "(" + item.Gender.ToString() + ")" + item.UserName.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).RichText.Substring(0, 3).SetFontColor(XLColor.Red).SetBold().SetStrikethrough();

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = "2022/05/18 13:27:12";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Style = dateStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = "1234560";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Style = numberStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = "123000";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Style = currencyStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).FormulaA1 = "=IF(" + BaseHelper.MergeNumberToChar(7, rowIndex) + "=\"F\",1,0)";
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //置中
                        worksheet.Columns("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
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
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

                        #region //隱藏
                        //worksheet.Column(1).Hide();
                        #endregion

                        #region //格式化
                        //string conditionalFormatRange = "G:G";
                        string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        {
                            range.AddConditionalFormat().WhenEquals("M")
                                //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                                .Font.SetFontColor(XLColor.BabyBlue);

                            range.AddConditionalFormat().WhenEquals("F")
                                //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                                .Font.SetFontColor(XLColor.Red);
                        }

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

        #region//ExcelGetOutsourcingDetail --托外單明細
        public void ExcelGetOutsourcingDetail(string OspStartDate = "", string OspEndDate = "", string OspCreatStartDate = "", string OspCreatEndDate = "", string OspNo = "", string SupplierNo = "", string SupplierShortName = "", string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = "", int ProcessAlias = -1, string BackSttus = "", string OsrErpFullNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetOutsourcingDetail(OspStartDate, OspEndDate, OspCreatStartDate, OspCreatEndDate, OspNo, SupplierNo, SupplierShortName
                    , WoErpFullNo, MtlItemNo, MtlItemName, ProcessAlias, BackSttus, OsrErpFullNo
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】托外單明細";
                    string excelsheetName = "托外單明細";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    List<OutsourcingResult> detailReportDatas = JsonConvert.DeserializeObject<List<OutsourcingResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "托外單號", "托外生產日期", "托外單創建日", "托外單創建者", "供應商編號", "供應商簡稱", "供應商全稱"
                        , "托外備註	", "製令", "產品品號", "產品品名", "規格", "托外生產數量", "供貨生產數量", "加工製程", "加工代碼", "單身添加日", "單身添加者"
                        , "預計回廠日", "托外進貨日", "托外進貨單", "托外進貨單拋轉狀態", "托外進貨單頭備註", "托外進貨品項序號", "托外進貨數", "托外進貨品項備註"
                        , "托外計時長(小時)", "判斷", "判斷備註" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 15)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            var receiptDate = item.ReceiptDate != null ? item.ReceiptDate.ToString() : "";
                            var osrFullErpNo = item.OsrErpNo != null ? item.OsrErpPrefix.ToString() + "-" + item.OsrErpNo.ToString() : "";

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.OspNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.OspDate.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.OspCreateDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.OspUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.SupplierNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.SupplierShortName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.SupplierName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.OspDesc.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.WoErpFullNo;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.MtlItemName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.MtlItemSpec;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.OspQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.SuppliedQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.ProcessAlias ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.ProcessCodeName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.OspDetailCreateDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.WoUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.ExpectedDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = receiptDate;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = osrFullErpNo;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.OpsTransferStatus.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.OspHRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.OsrSeq.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.ReceiptQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.OspBRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.OspTimes.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.BackStatus.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.BackStatusRemark.ToString();





                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 30;
                        //worksheet.Column(2).Width = 30;
                        //worksheet.Column(3).Width = 30;
                        //worksheet.Column(4).Width = 30;
                        //worksheet.Column(5).Width = 30;
                        //worksheet.Column(6).Width = 20;
                        //worksheet.Column(7).Width = 20;
                        //worksheet.Column(8).Width = 20;
                        //worksheet.Column(9).Width = 20;
                        //worksheet.Column(10).Width = 15;
                        //worksheet.Column(11).Width = 15;



                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelGeMoProgressDetail --製令加工進度 
        public void ExcelGeMoProgressDetail(string StartDate = "", string EndDate = "", string WoErpFullNo = "", string WoStatus = "", int ProdMode = -1, string MtlItemNo = ""
            , string MtlItemName = "", string MoStatus = "", string QuantityStatusQ = "", string ReceiptStatusQ = "", int MoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GeMoProgressDetail(StartDate, EndDate, WoErpFullNo, WoStatus, ProdMode, MtlItemNo, MtlItemName, MoStatus, QuantityStatusQ, ReceiptStatusQ, MoId
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】製令加工進度 ";
                    string excelsheetName = "製令加工進度 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "公司別", "工單狀態", "工單制單者", "ERP工單預計產量", "生產模式", "制令單號", "品號", "品名", "規格", "工單預計開工日"
                        , "工單預計完工日", "工單實際開工日", "工單實際完工日", "制令開單日", "制令條碼", "制令狀態", "制令制單者", "MES制令預計產量", "制令已領套數", "領料狀態"
                        , "制令總工站數", "加工進度(%)", "已入庫數", "制令已入庫狀態","制程","未开工","加工中","已完工","良品","不良","报废","良率" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            var ActualStart = item.ActualStart != null ? item.ActualStart.ToString() : "";
                            var ActualEnd = item.ActualEnd != null ? item.ActualEnd.ToString() : "";
                            var ExpectedStart = item.ExpectedStart != null ? item.ExpectedStart.ToString() : "";
                            var ExpectedEnd = item.ExpectedEnd != null ? item.ExpectedEnd.ToString() : "";

                            var ProcessProgress = item.ProcessProgress != null ? item.ProcessProgress : 0;
                            var ReceiptQtySum = item.ReceiptQtySum != null ? item.ReceiptQtySum : 0;
                            var ReceiptStatus = item.ReceiptStatus != null ? item.ReceiptStatus.ToString() : "";



                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.CompanyName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.WoStatus.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.WoCreater.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ERPPlanQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Mode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.WoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.MtlItemDesc;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = ExpectedStart;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = ExpectedEnd;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = ActualStart;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = ActualEnd;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.MoCreateDay.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.MoId.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.MoStatus.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.MoCreater.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.MoQuantity.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.RequisitionSetQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.QuantityStatus.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.MoTotalProcess.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = ProcessProgress.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = ReceiptQtySum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = ReceiptStatus.ToString();



                            if (item.ProcessDetail != null)
                            {
                                string ProcessDetailStr = item.ProcessDetail.ToString();
                                dynamic[] ProcessDetail = JsonConvert.DeserializeObject<dynamic[]>(JObject.Parse(ProcessDetailStr)["data"].ToString());

                                //var subitemrowIndex = rowIndex;


                                foreach (var subitem2 in ProcessDetail)
                                {

                                    var ProcessAlias = subitem2.ProcessAlias != null ? subitem2.ProcessAlias.ToString() : "";
                                    var NotStartQty = subitem2.NotStartQty != null ? subitem2.NotStartQty : 0;
                                    var StartQty = subitem2.StartQty != null ? subitem2.StartQty : 0;
                                    var WipQty = subitem2.WipQty != null ? subitem2.WipQty : 0;
                                    var TotalPassQty = subitem2.TotalPassQty != null ? subitem2.TotalPassQty : 0;
                                    var TotalNgQty = subitem2.TotalNgQty != null ? subitem2.TotalNgQty : 0;
                                    var TotalScrapQty = subitem2.TotalScrapQty != null ? subitem2.TotalScrapQty : 0;
                                    var PassRate = subitem2.PassRate != null ? subitem2.PassRate.ToString() : "";

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = ProcessAlias.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = NotStartQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = StartQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = WipQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = TotalPassQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = TotalNgQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = TotalScrapQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = PassRate.ToString();
                                    rowIndex++;

                                }
                                rowIndex--;
                            }
                            else
                            {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = "";
                                //rowIndex++;
                            }


                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                            worksheet.Range(startIndex, 5, rowIndex, 5).Merge();
                            worksheet.Range(startIndex, 6, rowIndex, 6).Merge();
                            worksheet.Range(startIndex, 7, rowIndex, 7).Merge();
                            worksheet.Range(startIndex, 8, rowIndex, 8).Merge();
                            worksheet.Range(startIndex, 9, rowIndex, 9).Merge();
                            worksheet.Range(startIndex, 10, rowIndex, 10).Merge();
                            worksheet.Range(startIndex, 11, rowIndex, 11).Merge();
                            worksheet.Range(startIndex, 12, rowIndex, 12).Merge();
                            worksheet.Range(startIndex, 13, rowIndex, 13).Merge();
                            worksheet.Range(startIndex, 14, rowIndex, 14).Merge();
                            worksheet.Range(startIndex, 15, rowIndex, 15).Merge();
                            worksheet.Range(startIndex, 16, rowIndex, 16).Merge();
                            worksheet.Range(startIndex, 17, rowIndex, 17).Merge();
                            worksheet.Range(startIndex, 18, rowIndex, 18).Merge();
                            worksheet.Range(startIndex, 19, rowIndex, 19).Merge();
                            worksheet.Range(startIndex, 20, rowIndex, 20).Merge();
                            worksheet.Range(startIndex, 21, rowIndex, 21).Merge();
                            worksheet.Range(startIndex, 22, rowIndex, 22).Merge();
                            worksheet.Range(startIndex, 23, rowIndex, 23).Merge();
                            worksheet.Range(startIndex, 24, rowIndex, 24).Merge();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 30;
                        //worksheet.Column(2).Width = 30;
                        //worksheet.Column(3).Width = 30;
                        //worksheet.Column(4).Width = 30;
                        //worksheet.Column(5).Width = 30;
                        //worksheet.Column(6).Width = 20;
                        //worksheet.Column(7).Width = 20;
                        //worksheet.Column(8).Width = 20;
                        //worksheet.Column(9).Width = 20;
                        //worksheet.Column(10).Width = 15;
                        //worksheet.Column(11).Width = 15;



                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //DownloadMeasureDataExcel 量測資料EXCEL下載
        public void DownloadMeasureDataExcel(int QcRecordId = -1, string BarcodeNo = "", string QcResult = "", string QcItemName = "", string MachineNumber = "", string MtlItemNo = "", string MtlItemName = "")
        {
            try
            {
                WebApiLoginCheck("MeasurementRecord", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcMeasureData(QcRecordId, BarcodeNo, QcResult, QcItemName, MachineNumber);
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
                    headerStyle.Fill.BackgroundColor = XLColor.AppleGreen;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
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
                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "量測數據_品名_" + MtlItemName;
                    string excelsheetName = "量測單號_" + QcRecordId.ToString();

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"]["result"].ToString());
                    string[] valueList = JObject.Parse(dataRequest)["data"]["valueList"].ToString().Replace("\r\n", "").Replace("[", "").Replace("]", "").Replace("\"", "").Replace(" ", "").Split(',');
                    string[] cellList = JObject.Parse(dataRequest)["data"]["cellList"].ToString().Replace("\r\n", "").Replace("[", "").Replace("]", "").Replace("\"", "").Replace(" ", "").Split(',');
                    string[] headerValues = new string[] { "序號", "球標", "檢測項目", "檢測備註", "設計值", "上公差", "下公差", "單位", "量測設備", "Z軸", "量測人員" };
                    string colIndexValue = "";
                    int rowIndex = 1;
                    int count = 0;
                    int colIndex = 0;

                    foreach(var item in cellList)
                    {
                        string cell = item.ToString();
                        if (cell.IndexOf('2') != -1)
                            break;
                        count++;
                    }

                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 4)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 8).Merge().Style = titleStyle;

                        worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = "Tray 條碼:";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = JObject.Parse(dataRequest)["data"]["QcRecordBarcodeNo"].ToString();
                        worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = "站別:";
                        worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = JObject.Parse(dataRequest)["data"]["QcRecordProcessAlias"].ToString();

                        rowIndex++;
                        #endregion

                        #region //HEADER
                        //for (int i = 0; i < headerValues.Length; i++)
                        //{
                        //    colIndexValue = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                        //    worksheet.Cell(colIndexValue).Value = headerValues[i];
                        //    worksheet.Cell(colIndexValue).Style = headerStyle;
                        //}
                        #endregion

                        #region //BODY
                        for (int i = 0; i < valueList.Length; i++)
                        {
                            if (colIndex == count)
                            {
                                rowIndex++;
                                colIndex = 0;
                            }
                            colIndexValue = BaseHelper.MergeNumberToChar(colIndex + 1, rowIndex);
                            if (rowIndex == 2) worksheet.Cell(colIndexValue).Style = headerStyle;
                            worksheet.Cell(colIndexValue).Value = valueList[i];
                            colIndex++;
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //置中
                        worksheet.Columns("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
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
                        worksheet.SheetView.FreezeRows(1);
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

        #region//OrderStockReportExcel 訂單進度&庫存數量統計報表

        public void OrderStockReportExcel(string LensModel,string LensModelName, string Unshipped
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetOrderStockReport(LensModel, LensModelName, Unshipped
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】訂單進度&庫存數量統計報表";
                    string excelsheetName = "訂單進度&庫存數量統計報表";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "鏡頭機種", "訂單總數", "已出貨數", "未出貨數", "主件庫存數", "元件料號", "元件庫存數", "製令未領數", "在製(已領未完工)", "已採未進貨" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;

                        foreach (var item in data)
                        {
                            rowIndex++;

                            startIndex = rowIndex;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.LensModel.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MainItemStock.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.OrderCount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ShippedQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.UnshippedQty.ToString();

                            if (item.ComponentData != null)
                            {
                                string componentDataStr = item.ComponentData.ToString();
                                dynamic[] componentData = JsonConvert.DeserializeObject<dynamic[]>(JObject.Parse(componentDataStr)["data"].ToString());

                                //var subitemrowIndex = rowIndex;


                                foreach (var subitem2 in componentData)
                                {

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = subitem2.ComponentMtlItemNo.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = subitem2.ComponentItemStock.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = subitem2.UnissuedQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = subitem2.RequisitionQty.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = subitem2.ReceivedNotStocked.ToString();
                                    rowIndex++;

                                }
                                rowIndex--;
                            }
                            else {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = "";
                                //rowIndex++;
                            }
                            

                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                            worksheet.Range(startIndex, 5, rowIndex, 5).Merge();
                            

                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;
                        #endregion

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

        #region//GetRptMoInformation2Excel 取得生產異常資料(業務用)

        public void GetRptMoInformation2Excel(string WoErpFullNo, string ModeDesc, string CustomerShortName, string SearchKey, string ExpectedStart, string ExpectedEnd
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRptMoInformation2(WoErpFullNo, ModeDesc, CustomerShortName, SearchKey, ExpectedStart, ExpectedEnd
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】生產進度異常查詢(業務用)";
                    string excelsheetName = "生產進度異常查詢(業務用)";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["result"].ToString());
                    string[] header = new string[] { "製令", "生產模式", "加工客戶", "品名", "品號", "最慢加工數量", "預計開工", "預計完工", "實際開工", "實際完工", "最慢加工站別" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;

                        foreach (var item in data)
                        {
                            rowIndex++;

                            startIndex = rowIndex;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.WoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ModeDesc.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.CustomerShortName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.LatestBarcodeQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ExpectedStart.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ExpectedEnd.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ActualStart.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ActualEnd.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.LatestProcess.ToString();
                           

                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;
                        #endregion

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

        #region//SendExcelPendingPurchaseRequestReport 寄送 已請未購明細(一般＆資產請購)
        [HttpPost]
        [Route("api/MES/apiSendExcelPendingPurchaseRequestReport")]
        public void SendExcelPendingPurchaseRequestReport(string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "SendExcelPendingPurchaseRequestReport");
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetPendingPurchaseRequestReport(Company);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    byte[] excelBytes;
                    string fileName = $"已請未購明細(一般＆資產請購)_{DateTime.Now:yyyyMMdd}.xlsx";

                    #region //樣式
                    // 設定樣式
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Font.FontName = "微軟正黑體";
                    defaultStyle.Font.FontSize = 12;
                    #endregion


                    #region //EXCEL

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("已請未購明細(一般＆資產請購)");
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        // 設定標題
                        worksheet.Cell("A1").Value = "已請未購明細(一般＆資產請購)";
                        var titleRange = worksheet.Range("A1:U1");
                        titleRange.Merge();
                        titleRange.Style.Font.FontSize = 16;
                        titleRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Row(1).Height = 40;

                        // 設定表頭
                        string[] headers = new string[] {
                    "請購日", "完成簽核時間", "單別", "請購單號",
                    "請購序號", "MES單號", "BPM單號", "品號", "品名", "規格",
                    "請購數量", "單位", "需用日", "請購人員", "請購部門",
                    "備註", "廠商代號", "廠商名稱", "採購人員", "庫別名稱"
                };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.FontSize = 14;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }

                        // 寫入數據
                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                        int rowIndex = 3;

                        #region //BODY

                        foreach (var item in data)
                        {
                            // 設定一般資料格式
                            var row = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                            row.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            worksheet.Cell(rowIndex, 1).Value = item.purchase_date?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 2).Value = item.approval_complete_time?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 3).Value = item.document_type?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 4).Value = item.purchase_order_no?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 5).Value = item.purchase_seq_no?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 6).Value = item.mes_no?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 7).Value = item.bpm_no?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 8).Value = item.item_no?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 9).Value = item.item_name?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 10).Value = item.specification?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 11).Value = item.purchase_qty?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 12).Value = item.unit?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 13).Value = item.required_date?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 14).Value = item.request_dept?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 15).Value = item.requester?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 16).Value = item.remarks?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 17).Value = item.vendor_code?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 18).Value = item.vendor_name?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 19).Value = item.purchaser?.ToString() ?? "";
                            worksheet.Cell(rowIndex, 20).Value = item.warehouse_name?.ToString() ?? "";

                            // 設定每個儲存格的邊框
                            for (int i = 1; i <= headers.Length; i++)
                            {
                                worksheet.Cell(rowIndex, i).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(rowIndex, i).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(rowIndex, i).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(rowIndex, i).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            }

                            rowIndex++;
                        }

                        // 設定工作表格式
                        worksheet.Columns().AdjustToContents();
                        worksheet.SheetView.FreezeRows(2);

                        #endregion

                        #region 輸出檔案
                        using (MemoryStream ms = new MemoryStream())
                        {
                            workbook.SaveAs(ms);
                            excelBytes = ms.ToArray();
                        }
                        #endregion
                    }
                    #endregion

                    #region//Mail 寄送已請未購明細

                    //取得Mail樣板設定
                    string SettingSchema = "PendingPurchaseRequestReport";
                    string SettingNo = "Y";

                    mesReportDA = new MesReportDA();
                    string mailSettingsJson = mesReportDA.SendQcMail(SettingSchema, SettingNo);
                    if (string.IsNullOrEmpty(mailSettingsJson))
                    {
                        throw new Exception("無法取得郵件設定");
                    }

                    var parsedSettings = JObject.Parse(mailSettingsJson);
                    if (parsedSettings["status"].ToString() != "success")
                    {
                        throw new Exception(parsedSettings["msg"].ToString());
                    }

                    dynamic mailSettings = parsedSettings["data"];

                    foreach (var item in mailSettings)
                    {
                        var mailFile = new MailFile
                        {
                            FileName = Path.GetFileNameWithoutExtension(fileName),
                            FileExtension = ".xlsx",
                            FileContent = excelBytes
                        };


                        #region //寄送Mail
                        var mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = $"已請未購明細(一般＆資產請購)_{DateTime.Now:yyyy/MM/dd}",
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = item.MailTo,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = $"附件為已請未購明細(一般＆資產請購)，日期:{DateTime.Now:yyyy/MM/dd}",
                            TextBody = "-",
                            FileInfo = new List<MailFile> { mailFile },
                            QcFileFlag = "N"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion

                    Response.Write(JObject.FromObject(new
                    {
                        status = "success",
                        msg = "報表已成功生成並寄送"
                    }).ToString());
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

        #region//InProcessGoodsDefectRateReportExcel 在製品不良率報表

        public void InProcessGoodsDefectRateReportExcel(int ModeId = -1, string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        {
            try
            {
                //WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetInProcessGoodsDefectRateReport(ModeId, MtlItemNo, StartDate, EndDate
                                                            , OrderBy, PageIndex, PageSize, RedirectType);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】在製品不良率報表";
                    string excelsheetName = "在製品不良率報表";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["result"].ToString());
                    string[] header = new string[] { "產品品號", "產品品名", "總投入數", "總完工數", "總報廢數", "日期", "投入數", "良品數", "不良品數", "良率(%)", "不良原因", "不良數" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 13).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;

                        foreach (var item in data)
                        {
                            rowIndex++;

                            startIndex = rowIndex;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.sumInputQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.sumCompleteQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.sumScrapQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.FinishDate01.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.sumStationQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.sumPassQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.sumNgQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.PassRate.ToString();

                            if (item.InvalidData != null)
                            {
                                string InvalidDataStr = item.InvalidData.ToString();
                                dynamic[] InvalidData = JsonConvert.DeserializeObject<dynamic[]>(JObject.Parse(InvalidDataStr)["data"].ToString());

                                //var subitemrowIndex = rowIndex;


                                foreach (var subitem2 in InvalidData)
                                {

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = subitem2.CauseDesc.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = subitem2.sumNgQty.ToString();
                                    rowIndex++;

                                }
                                rowIndex--;
                            }
                            else
                            {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = "";
                                //rowIndex++;
                            }


                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                            worksheet.Range(startIndex, 5, rowIndex, 5).Merge();
                            worksheet.Range(startIndex, 6, rowIndex, 6).Merge();
                            worksheet.Range(startIndex, 7, rowIndex, 7).Merge();
                            worksheet.Range(startIndex, 8, rowIndex, 8).Merge();
                            worksheet.Range(startIndex, 9, rowIndex, 9).Merge();
                            worksheet.Range(startIndex, 10, rowIndex, 10).Merge();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelProjectOutCycle 托外進貨週期性報表
        public void ExcelProjectOutCycle(string Company = "", string StartReceiptDate = "", string EndReceiptDate = "", string ConfirmStatusProjectOutCycle = "", string CheckoutStatusProjectOutCycle = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProjectOutCycle(Company, StartReceiptDate.Replace("-", ""), EndReceiptDate.Replace("-", ""),
                    ConfirmStatusProjectOutCycle, CheckoutStatusProjectOutCycle);
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
                    defaultStyle.Font.FontName = "微軟正黑體";
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
                    titleStyle.Font.FontName = "微軟正黑體";
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
                    headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                    headerStyle.Font.FontName = "微軟正黑體";
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dataStyle
                    var dataStyle = XLWorkbook.DefaultStyle;
                    dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.TopBorderColor = XLColor.Black;
                    dataStyle.Border.BottomBorderColor = XLColor.Black;
                    dataStyle.Border.LeftBorderColor = XLColor.Black;
                    dataStyle.Border.RightBorderColor = XLColor.Black;
                    dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dataStyle.Font.FontName = "微軟正黑體";
                    dataStyle.Font.FontSize = 12;
                    dataStyle.Font.Bold = false;
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
                    numberStyle.Font.FontName = "微軟正黑體";
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0.00";
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
                    currencyStyle.Font.FontName = "微軟正黑體";
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "¥#,##0.00";
                    #endregion
                    #endregion

                    #region //EXCEL

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("托外進貨週期性報表");
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = "托外進貨週期性報表";
                        var titleRange = worksheet.Range("A1:K1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        // 設定表頭
                        string[] headers = new string[] {
                    "進貨日期", "託外進貨單號", "廠別名稱", "加工廠商代號", "加工廠商簡稱",
                    "廠商單號", "發票聯數", "課稅別", "發票號碼", "付款條件代號", "付款條件名稱",
                    "件數", "幣別", "匯率", "原幣扣款金額", "本幣進貨費用",
                    "數量合計", "包裝數量合計", "原未稅款金額", "原幣稅額", "原幣金額合計",
                    "本幣未稅金額", "本幣稅額", "本幣金額合計", "品號", "品名",
                    "規格", "驗收日期", "庫別", "儲位", "製程代號", "製程名稱", "進貨數量",
                    "驗收數量","有效日期","進貨包裝數量","驗收包裝數量","複檢日期","報廢數量",
                    "驗退數量","計價數量","報廢包裝數量","驗退包裝數量","包裝單位","單位","計價單位",
                    "扣款說明","原幣加工單價","原幣扣款金額","應付工繳","本幣進貨費用","原未稅款金額",
                    "原幣稅額","原幣金額合計","本幣未稅金額","本幣稅額","本幣金額合計","逾期","檢驗狀態",
                    "檢退","製令單號","備註","批號","急料","結帳","儲存位置","應付憑單號"

                };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style = headerStyle;
                        }

                        // 寫入數據
                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                        int rowIndex = 3;

                        #region //BODY

                        foreach (var item in data)
                        {
                            // 設定一般資料格式
                            var row = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                            row.Style = dataStyle;

                            worksheet.Cell(rowIndex, 1).Value = item.TH003.ToString();
                            worksheet.Cell(rowIndex, 2).Value = item.TH00A.ToString();
                            worksheet.Cell(rowIndex, 3).Value = item.MB002.ToString();
                            worksheet.Cell(rowIndex, 4).Value = item.TH005.ToString();
                            worksheet.Cell(rowIndex, 5).Value = item.MA002.ToString();
                            worksheet.Cell(rowIndex, 6).Value = item.TH006.ToString();
                            worksheet.Cell(rowIndex, 7).Value = item.NM002.ToString();
                            worksheet.Cell(rowIndex, 8).Value = item.TypeName.ToString();
                            worksheet.Cell(rowIndex, 9).Value = item.TH014.ToString();
                            worksheet.Cell(rowIndex, 10).Value = item.TH033.ToString();
                            worksheet.Cell(rowIndex, 11).Value = item.NA003.ToString();
                            worksheet.Cell(rowIndex, 12).Value = item.TH009.ToString();
                            //worksheet.Cell(rowIndex, 13).Value = item.TH007.ToString();
                            //worksheet.Cell(rowIndex, 14).Value = item.TH008.ToString();
                            //worksheet.Cell(rowIndex, 15).Value = item.TH019.ToString();
                            //worksheet.Cell(rowIndex, 16).Value = item.TH021.ToString();
                            worksheet.Cell(rowIndex, 17).Value = item.TH022.ToString();
                            worksheet.Cell(rowIndex, 18).Value = item.TH034.ToString();
                            //worksheet.Cell(rowIndex, 19).Value = item.TH027.ToString();
                            //worksheet.Cell(rowIndex, 20).Value = item.TH020.ToString();
                            //worksheet.Cell(rowIndex, 21).Value = item.TH02027.ToString();
                            //worksheet.Cell(rowIndex, 22).Value = item.TH031.ToString();
                            //worksheet.Cell(rowIndex, 23).Value = item.TH032.ToString();
                            //worksheet.Cell(rowIndex, 24).Value = item.TH03132.ToString();
                            worksheet.Cell(rowIndex, 25).Value = item.TI004.ToString();
                            worksheet.Cell(rowIndex, 26).Value = item.TI005.ToString();
                            worksheet.Cell(rowIndex, 27).Value = item.TI006.ToString();
                            worksheet.Cell(rowIndex, 28).Value = item.TI018.ToString();
                            worksheet.Cell(rowIndex, 29).Value = item.MC002.ToString();
                            worksheet.Cell(rowIndex, 30).Value = item.TI057.ToString();
                            worksheet.Cell(rowIndex, 31).Value = item.TI015.ToString();
                            worksheet.Cell(rowIndex, 32).Value = item.MW002.ToString();
                            worksheet.Cell(rowIndex, 33).Value = item.TI007.ToString();
                            worksheet.Cell(rowIndex, 34).Value = item.TI019.ToString();
                            worksheet.Cell(rowIndex, 35).Value = item.TI011.ToString();
                            worksheet.Cell(rowIndex, 36).Value = item.TI016.ToString();
                            worksheet.Cell(rowIndex, 37).Value = item.TI017.ToString();
                            worksheet.Cell(rowIndex, 38).Value = item.TI012.ToString();
                            worksheet.Cell(rowIndex, 39).Value = item.TI021.ToString();
                            worksheet.Cell(rowIndex, 40).Value = item.TI022.ToString();
                            worksheet.Cell(rowIndex, 41).Value = item.TI020.ToString();
                            worksheet.Cell(rowIndex, 42).Value = item.TI050.ToString();
                            worksheet.Cell(rowIndex, 43).Value = item.TI051.ToString();
                            worksheet.Cell(rowIndex, 44).Value = item.TI049.ToString();
                            worksheet.Cell(rowIndex, 45).Value = item.TI008.ToString();
                            worksheet.Cell(rowIndex, 46).Value = item.TI023.ToString();
                            worksheet.Cell(rowIndex, 47).Value = item.TI028.ToString();
                            //worksheet.Cell(rowIndex, 48).Value = item.TI024.ToString();
                            //worksheet.Cell(rowIndex, 49).Value = item.TI026.ToString();
                            //worksheet.Cell(rowIndex, 50).Value = item.TI02420TH008.ToString();
                            //worksheet.Cell(rowIndex, 51).Value = item.TI04647.ToString();
                            //worksheet.Cell(rowIndex, 52).Value = item.TI044.ToString();
                            //worksheet.Cell(rowIndex, 53).Value = item.TI045.ToString();
                            //worksheet.Cell(rowIndex, 54).Value = item.TI0442045.ToString();
                            //worksheet.Cell(rowIndex, 55).Value = item.TI046.ToString();
                            //worksheet.Cell(rowIndex, 56).Value = item.TI047.ToString();
                            //worksheet.Cell(rowIndex, 57).Value = item.TI0462047.ToString();
                            worksheet.Cell(rowIndex, 58).Value = item.TI034.ToString();
                            worksheet.Cell(rowIndex, 59).Value = item.TI035Status.ToString();
                            worksheet.Cell(rowIndex, 60).Value = item.TI036.ToString();
                            worksheet.Cell(rowIndex, 61).Value = item.TI01314.ToString();
                            worksheet.Cell(rowIndex, 62).Value = item.TI040.ToString();
                            worksheet.Cell(rowIndex, 63).Value = item.TI010.ToString();
                            worksheet.Cell(rowIndex, 64).Value = item.TI048.ToString();
                            worksheet.Cell(rowIndex, 65).Value = item.TI038.ToString();
                            worksheet.Cell(rowIndex, 66).Value = item.MC003.ToString();
                            worksheet.Cell(rowIndex, 67).Value = item.TI02930.ToString();

                            // 設定數字格式
                            //worksheet.Range(rowIndex, 14, rowIndex, 16).Style = numberStyle;
                            //worksheet.Range(rowIndex, 19, rowIndex, 24).Style = currencyStyle;

                            rowIndex++;
                        }
                        // 設定小計行
                        worksheet.Cell(rowIndex, 1).Value = "小計：" + data.Count().ToString();
                        var subtotalRow = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                        subtotalRow.Style = headerStyle;

                        // 計算小計
                        // 本幣進貨費用
                        decimal subtotalTH021 = data.Sum(x => decimal.TryParse(x.TH021?.ToString(), out decimal val) ? val : 0);
                        // 數量合計
                        decimal subtotalTH022 = data.Sum(x => decimal.TryParse(x.TH022?.ToString(), out decimal val) ? val : 0);
                        // 包裝數量合計
                        decimal subtotalTH034 = data.Sum(x => decimal.TryParse(x.TH034?.ToString(), out decimal val) ? val : 0);
                        // 本幣未稅金額
                        decimal subtotalTH031 = data.Sum(x => decimal.TryParse(x.TH031?.ToString(), out decimal val) ? val : 0);
                        // 本幣稅額
                        decimal subtotalTH032 = data.Sum(x => decimal.TryParse(x.TH032?.ToString(), out decimal val) ? val : 0);
                        // 本幣金額合計
                        decimal subtotalTH03132 = data.Sum(x => decimal.TryParse(x.TH03132?.ToString(), out decimal val) ? val : 0);
                        // 進貨數量
                        decimal subtotalTI007 = data.Sum(x => decimal.TryParse(x.TI007?.ToString(), out decimal val) ? val : 0);
                        // 驗收數量
                        decimal subtotalTI019 = data.Sum(x => decimal.TryParse(x.TI019?.ToString(), out decimal val) ? val : 0);
                        // 進貨包裝數量
                        decimal subtotalTI016 = data.Sum(x => decimal.TryParse(x.TI016?.ToString(), out decimal val) ? val : 0);
                        // 驗收包裝數量
                        decimal subtotalTI017 = data.Sum(x => decimal.TryParse(x.TI017?.ToString(), out decimal val) ? val : 0);
                        // 報廢數量
                        decimal subtotalTI021 = data.Sum(x => decimal.TryParse(x.TI021?.ToString(), out decimal val) ? val : 0);
                        // 驗退數量
                        decimal subtotalTI022 = data.Sum(x => decimal.TryParse(x.TI022?.ToString(), out decimal val) ? val : 0);
                        // 計價數量
                        decimal subtotalTI020 = data.Sum(x => decimal.TryParse(x.TI020?.ToString(), out decimal val) ? val : 0);
                        // 報廢包裝數量
                        decimal subtotalTI050 = data.Sum(x => decimal.TryParse(x.TI050?.ToString(), out decimal val) ? val : 0);
                        // 驗退包裝數量
                        decimal subtotalTI051 = data.Sum(x => decimal.TryParse(x.TI051?.ToString(), out decimal val) ? val : 0);
                        // 原幣扣款金額
                        decimal subtotalTI02420TH008 = data.Sum(x => decimal.TryParse(x.TI02420TH008?.ToString(), out decimal val) ? val : 0);

                        // 設定小計欄位值
                        //worksheet.Cell(rowIndex, 16).Value = subtotalTH021;  // 本幣進貨費用
                        worksheet.Cell(rowIndex, 17).Value = subtotalTH022;  // 數量合計
                        worksheet.Cell(rowIndex, 18).Value = subtotalTH034;  // 包裝數量合計
                                                                             //worksheet.Cell(rowIndex, 22).Value = subtotalTH031;  // 本幣未稅金額
                                                                             //worksheet.Cell(rowIndex, 23).Value = subtotalTH032;  // 本幣稅額
                                                                             //worksheet.Cell(rowIndex, 24).Value = subtotalTH03132;  // 本幣金額合計

                        // 根據Excel中的位置設定其他欄位

                        worksheet.Cell(rowIndex, 33).Value = subtotalTI007;  // 進貨數量
                        worksheet.Cell(rowIndex, 34).Value = subtotalTI019;  // 驗收數量
                        worksheet.Cell(rowIndex, 36).Value = subtotalTI016;  // 進貨包裝數量
                        worksheet.Cell(rowIndex, 37).Value = subtotalTI017;  // 驗收包裝數量
                        worksheet.Cell(rowIndex, 39).Value = subtotalTI021;  // 報廢數量
                        worksheet.Cell(rowIndex, 40).Value = subtotalTI022;  // 驗退數量
                        worksheet.Cell(rowIndex, 41).Value = subtotalTI020;  // 計價數量
                        worksheet.Cell(rowIndex, 42).Value = subtotalTI050;  // 報廢包裝數量
                        worksheet.Cell(rowIndex, 43).Value = subtotalTI051;  // 驗退包裝數量
                                                                             //worksheet.Cell(rowIndex, 49).Value = subtotalTI02420TH008;  // 原幣扣款金額
                                                                             //worksheet.Cell(rowIndex, 51).Value = subtotalTI02420TH008;  // 本幣進貨費用
                                                                             // 設定小計行的數字格式
                        var subtotalRowRange = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                        subtotalRowRange.Style.NumberFormat.Format = "#,##0.00";

                        // 設定工作表格式
                        worksheet.Columns().AdjustToContents();
                        worksheet.SheetView.FreezeRows(2);

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

                        #endregion

                        #region 輸出檔案
                        using (MemoryStream output = new MemoryStream())
                        {
                            workbook.SaveAs(output);
                            string fileGuid = Guid.NewGuid().ToString();
                            Session[fileGuid] = output.ToArray();

                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "OK",
                                fileGuid,
                                fileName = "托外進貨週期性報表",
                                fileExtension = ".xlsx"
                            });
                        }
                        #endregion
                    }
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

        #region//SendExcelProjectOutCycle 寄送托外進貨週期性報表
        [HttpPost]
        [Route("api/MES/apiSendExcelProjectOutCycle")]
        public void SendExcelProjectOutCycle(string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "SendExcelProjectOutCycle");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyyMMdd");
                string StartReceiptDate = today;
                string EndReceiptDate = today;
                string ConfirmStatusProjectOutCycle = "Y";
                string CheckoutStatusProjectOutCycle = "-1";
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProjectOutCycle(Company, StartReceiptDate.Replace("-", ""), EndReceiptDate.Replace("-", ""),
                    ConfirmStatusProjectOutCycle, CheckoutStatusProjectOutCycle);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    byte[] excelBytes;
                    string fileName = $"托外進貨週期性報表_{DateTime.Now:yyyyMMdd}.xlsx";

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
                    defaultStyle.Font.FontName = "微軟正黑體";
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
                    titleStyle.Font.FontName = "微軟正黑體";
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
                    headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                    headerStyle.Font.FontName = "微軟正黑體";
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dataStyle
                    var dataStyle = XLWorkbook.DefaultStyle;
                    dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    dataStyle.Border.TopBorderColor = XLColor.Black;
                    dataStyle.Border.BottomBorderColor = XLColor.Black;
                    dataStyle.Border.LeftBorderColor = XLColor.Black;
                    dataStyle.Border.RightBorderColor = XLColor.Black;
                    dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                    dataStyle.Font.FontName = "微軟正黑體";
                    dataStyle.Font.FontSize = 12;
                    dataStyle.Font.Bold = false;
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
                    numberStyle.Font.FontName = "微軟正黑體";
                    numberStyle.Font.FontSize = 12;
                    numberStyle.Font.Bold = false;
                    numberStyle.NumberFormat.Format = "#,##0.00";
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
                    currencyStyle.Font.FontName = "微軟正黑體";
                    currencyStyle.Font.FontSize = 12;
                    currencyStyle.Font.Bold = false;
                    currencyStyle.NumberFormat.Format = "#,##0.00";
                    #endregion
                    #endregion

                    #region //EXCEL

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("托外進貨週期性報表");
                        worksheet.Style = defaultStyle;

                        // 設定標題
                        worksheet.Cell("A1").Value = "托外進貨週期性報表";
                        var titleRange = worksheet.Range("A1:K1");
                        titleRange.Merge().Style = titleStyle;
                        worksheet.Row(1).Height = 40;

                        // 設定表頭
                        string[] headers = new string[] {
                    "進貨日期", "託外進貨單號", "廠別名稱", "加工廠商代號", "加工廠商簡稱",
                    "廠商單號", "發票聯數", "課稅別", "發票號碼", "付款條件代號", "付款條件名稱",
                    "件數", "幣別", "匯率", "原幣扣款金額", "本幣進貨費用",
                    "數量合計", "包裝數量合計", "原未稅款金額", "原幣稅額", "原幣金額合計",
                    "本幣未稅金額", "本幣稅額", "本幣金額合計", "品號", "品名",
                    "規格", "驗收日期", "庫別", "儲位", "製程代號", "製程名稱", "進貨數量",
                    "驗收數量","有效日期","進貨包裝數量","驗收包裝數量","複檢日期","報廢數量",
                    "驗退數量","計價數量","報廢包裝數量","驗退包裝數量","包裝單位","單位","計價單位",
                    "扣款說明","原幣加工單價","原幣扣款金額","應付工繳","本幣進貨費用","原未稅款金額",
                    "原幣稅額","原幣金額合計","本幣未稅金額","本幣稅額","本幣金額合計","逾期","檢驗狀態",
                    "檢退","製令單號","備註","批號","急料","結帳","儲存位置","應付憑單號"

                };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = worksheet.Cell(2, i + 1);
                            cell.Value = headers[i];
                            cell.Style = headerStyle;
                        }

                        // 寫入數據
                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                        int rowIndex = 3;

                        #region //BODY

                        foreach (var item in data)
                        {
                            // 設定一般資料格式
                            var row = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                            row.Style = dataStyle;

                            worksheet.Cell(rowIndex, 1).Value = item.TH003.ToString();
                            worksheet.Cell(rowIndex, 2).Value = item.TH00A.ToString();
                            worksheet.Cell(rowIndex, 3).Value = item.MB002.ToString();
                            worksheet.Cell(rowIndex, 4).Value = item.TH005.ToString();
                            worksheet.Cell(rowIndex, 5).Value = item.MA002.ToString();
                            worksheet.Cell(rowIndex, 6).Value = item.TH006.ToString();
                            worksheet.Cell(rowIndex, 7).Value = item.NM002.ToString();
                            worksheet.Cell(rowIndex, 8).Value = item.TypeName.ToString();
                            worksheet.Cell(rowIndex, 9).Value = item.TH014.ToString();
                            worksheet.Cell(rowIndex, 10).Value = item.TH033.ToString();
                            worksheet.Cell(rowIndex, 11).Value = item.NA003.ToString();
                            worksheet.Cell(rowIndex, 12).Value = item.TH009.ToString();
                            //worksheet.Cell(rowIndex, 13).Value = item.TH007.ToString();
                            //worksheet.Cell(rowIndex, 14).Value = item.TH008.ToString();
                            //worksheet.Cell(rowIndex, 15).Value = item.TH019.ToString();
                            //worksheet.Cell(rowIndex, 16).Value = item.TH021.ToString();
                            worksheet.Cell(rowIndex, 17).Value = item.TH022.ToString();
                            worksheet.Cell(rowIndex, 18).Value = item.TH034.ToString();
                            //worksheet.Cell(rowIndex, 19).Value = item.TH027.ToString();
                            //worksheet.Cell(rowIndex, 20).Value = item.TH020.ToString();
                            //worksheet.Cell(rowIndex, 21).Value = item.TH02027.ToString();
                            //worksheet.Cell(rowIndex, 22).Value = item.TH031.ToString();
                            //worksheet.Cell(rowIndex, 23).Value = item.TH032.ToString();
                            //worksheet.Cell(rowIndex, 24).Value = item.TH03132.ToString();
                            worksheet.Cell(rowIndex, 25).Value = item.TI004.ToString();
                            worksheet.Cell(rowIndex, 26).Value = item.TI005.ToString();
                            worksheet.Cell(rowIndex, 27).Value = item.TI006.ToString();
                            worksheet.Cell(rowIndex, 28).Value = item.TI018.ToString();
                            worksheet.Cell(rowIndex, 29).Value = item.MC002.ToString();
                            worksheet.Cell(rowIndex, 30).Value = item.TI057.ToString();
                            worksheet.Cell(rowIndex, 31).Value = item.TI015.ToString();
                            worksheet.Cell(rowIndex, 32).Value = item.MW002.ToString();
                            worksheet.Cell(rowIndex, 33).Value = item.TI007.ToString();
                            worksheet.Cell(rowIndex, 34).Value = item.TI019.ToString();
                            worksheet.Cell(rowIndex, 35).Value = item.TI011.ToString();
                            worksheet.Cell(rowIndex, 36).Value = item.TI016.ToString();
                            worksheet.Cell(rowIndex, 37).Value = item.TI017.ToString();
                            worksheet.Cell(rowIndex, 38).Value = item.TI012.ToString();
                            worksheet.Cell(rowIndex, 39).Value = item.TI021.ToString();
                            worksheet.Cell(rowIndex, 40).Value = item.TI022.ToString();
                            worksheet.Cell(rowIndex, 41).Value = item.TI020.ToString();
                            worksheet.Cell(rowIndex, 42).Value = item.TI050.ToString();
                            worksheet.Cell(rowIndex, 43).Value = item.TI051.ToString();
                            worksheet.Cell(rowIndex, 44).Value = item.TI049.ToString();
                            worksheet.Cell(rowIndex, 45).Value = item.TI008.ToString();
                            worksheet.Cell(rowIndex, 46).Value = item.TI023.ToString();
                            worksheet.Cell(rowIndex, 47).Value = item.TI028.ToString();
                            //worksheet.Cell(rowIndex, 48).Value = item.TI024.ToString();
                            //worksheet.Cell(rowIndex, 49).Value = item.TI026.ToString();
                            //worksheet.Cell(rowIndex, 50).Value = item.TI02420TH008.ToString();
                            //worksheet.Cell(rowIndex, 51).Value = item.TI04647.ToString();
                            //worksheet.Cell(rowIndex, 52).Value = item.TI044.ToString();
                            //worksheet.Cell(rowIndex, 53).Value = item.TI045.ToString();
                            //worksheet.Cell(rowIndex, 54).Value = item.TI0442045.ToString();
                            //worksheet.Cell(rowIndex, 55).Value = item.TI046.ToString();
                            //worksheet.Cell(rowIndex, 56).Value = item.TI047.ToString();
                            //worksheet.Cell(rowIndex, 57).Value = item.TI0462047.ToString();
                            worksheet.Cell(rowIndex, 58).Value = item.TI034.ToString();
                            worksheet.Cell(rowIndex, 59).Value = item.TI035Status.ToString();
                            worksheet.Cell(rowIndex, 60).Value = item.TI036.ToString();
                            worksheet.Cell(rowIndex, 61).Value = item.TI01314.ToString();
                            worksheet.Cell(rowIndex, 62).Value = item.TI040.ToString();
                            worksheet.Cell(rowIndex, 63).Value = item.TI010.ToString();
                            worksheet.Cell(rowIndex, 64).Value = item.TI048.ToString();
                            worksheet.Cell(rowIndex, 65).Value = item.TI038.ToString();
                            worksheet.Cell(rowIndex, 66).Value = item.MC003.ToString();
                            worksheet.Cell(rowIndex, 67).Value = item.TI02930.ToString();

                            // 設定數字格式
                            //worksheet.Range(rowIndex, 14, rowIndex, 16).Style = numberStyle;
                            //worksheet.Range(rowIndex, 19, rowIndex, 24).Style = currencyStyle;

                            rowIndex++;
                        }
                        // 設定小計行
                        worksheet.Cell(rowIndex, 1).Value = "小計：" + data.Count().ToString();
                        var subtotalRow = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                        subtotalRow.Style = headerStyle;

                        // 計算小計
                        // 本幣進貨費用
                        decimal subtotalTH021 = data.Sum(x => decimal.TryParse(x.TH021?.ToString(), out decimal val) ? val : 0);
                        // 數量合計
                        decimal subtotalTH022 = data.Sum(x => decimal.TryParse(x.TH022?.ToString(), out decimal val) ? val : 0);
                        // 包裝數量合計
                        decimal subtotalTH034 = data.Sum(x => decimal.TryParse(x.TH034?.ToString(), out decimal val) ? val : 0);
                        // 本幣未稅金額
                        decimal subtotalTH031 = data.Sum(x => decimal.TryParse(x.TH031?.ToString(), out decimal val) ? val : 0);
                        // 本幣稅額
                        decimal subtotalTH032 = data.Sum(x => decimal.TryParse(x.TH032?.ToString(), out decimal val) ? val : 0);
                        // 本幣金額合計
                        decimal subtotalTH03132 = data.Sum(x => decimal.TryParse(x.TH03132?.ToString(), out decimal val) ? val : 0);
                        // 進貨數量
                        decimal subtotalTI007 = data.Sum(x => decimal.TryParse(x.TI007?.ToString(), out decimal val) ? val : 0);
                        // 驗收數量
                        decimal subtotalTI019 = data.Sum(x => decimal.TryParse(x.TI019?.ToString(), out decimal val) ? val : 0);
                        // 進貨包裝數量
                        decimal subtotalTI016 = data.Sum(x => decimal.TryParse(x.TI016?.ToString(), out decimal val) ? val : 0);
                        // 驗收包裝數量
                        decimal subtotalTI017 = data.Sum(x => decimal.TryParse(x.TI017?.ToString(), out decimal val) ? val : 0);
                        // 報廢數量
                        decimal subtotalTI021 = data.Sum(x => decimal.TryParse(x.TI021?.ToString(), out decimal val) ? val : 0);
                        // 驗退數量
                        decimal subtotalTI022 = data.Sum(x => decimal.TryParse(x.TI022?.ToString(), out decimal val) ? val : 0);
                        // 計價數量
                        decimal subtotalTI020 = data.Sum(x => decimal.TryParse(x.TI020?.ToString(), out decimal val) ? val : 0);
                        // 報廢包裝數量
                        decimal subtotalTI050 = data.Sum(x => decimal.TryParse(x.TI050?.ToString(), out decimal val) ? val : 0);
                        // 驗退包裝數量
                        decimal subtotalTI051 = data.Sum(x => decimal.TryParse(x.TI051?.ToString(), out decimal val) ? val : 0);
                        // 原幣扣款金額
                        decimal subtotalTI02420TH008 = data.Sum(x => decimal.TryParse(x.TI02420TH008?.ToString(), out decimal val) ? val : 0);

                        // 設定小計欄位值
                        //worksheet.Cell(rowIndex, 16).Value = subtotalTH021;  // 本幣進貨費用
                        worksheet.Cell(rowIndex, 17).Value = subtotalTH022;  // 數量合計
                        worksheet.Cell(rowIndex, 18).Value = subtotalTH034;  // 包裝數量合計
                                                                             //worksheet.Cell(rowIndex, 22).Value = subtotalTH031;  // 本幣未稅金額
                                                                             //worksheet.Cell(rowIndex, 23).Value = subtotalTH032;  // 本幣稅額
                                                                             //worksheet.Cell(rowIndex, 24).Value = subtotalTH03132;  // 本幣金額合計

                        // 根據Excel中的位置設定其他欄位

                        worksheet.Cell(rowIndex, 33).Value = subtotalTI007;  // 進貨數量
                        worksheet.Cell(rowIndex, 34).Value = subtotalTI019;  // 驗收數量
                        worksheet.Cell(rowIndex, 36).Value = subtotalTI016;  // 進貨包裝數量
                        worksheet.Cell(rowIndex, 37).Value = subtotalTI017;  // 驗收包裝數量
                        worksheet.Cell(rowIndex, 39).Value = subtotalTI021;  // 報廢數量
                        worksheet.Cell(rowIndex, 40).Value = subtotalTI022;  // 驗退數量
                        worksheet.Cell(rowIndex, 41).Value = subtotalTI020;  // 計價數量
                        worksheet.Cell(rowIndex, 42).Value = subtotalTI050;  // 報廢包裝數量
                        worksheet.Cell(rowIndex, 43).Value = subtotalTI051;  // 驗退包裝數量
                                                                             //worksheet.Cell(rowIndex, 49).Value = subtotalTI02420TH008;  // 原幣扣款金額
                                                                             //worksheet.Cell(rowIndex, 51).Value = subtotalTI02420TH008;  // 本幣進貨費用
                                                                             // 設定小計行的數字格式
                        var subtotalRowRange = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                        subtotalRowRange.Style.NumberFormat.Format = "#,##0.00";

                        // 設定工作表格式
                        worksheet.Columns().AdjustToContents();
                        worksheet.SheetView.FreezeRows(2);

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

                        #endregion

                        #region 輸出檔案
                        using (MemoryStream ms = new MemoryStream())
                        {
                            workbook.SaveAs(ms);
                            excelBytes = ms.ToArray();
                        }
                        #endregion
                    }
                    #endregion

                    #region//Mail 寄送托外進貨週期性報表

                    //取得Mail樣板設定
                    string SettingSchema = "ProjectOutCycle";
                    string SettingNo = "Y";

                    mesReportDA = new MesReportDA();
                    string mailSettingsJson = mesReportDA.SendQcMail(SettingSchema, SettingNo);
                    if (string.IsNullOrEmpty(mailSettingsJson))
                    {
                        throw new Exception("無法取得郵件設定");
                    }

                    var parsedSettings = JObject.Parse(mailSettingsJson);
                    if (parsedSettings["status"].ToString() != "success")
                    {
                        throw new Exception(parsedSettings["msg"].ToString());
                    }

                    dynamic mailSettings = parsedSettings["data"];

                    foreach (var item in mailSettings)
                    {
                        var mailFile = new MailFile
                        {
                            FileName = Path.GetFileNameWithoutExtension(fileName),
                            FileExtension = ".xlsx",
                            FileContent = excelBytes
                        };


                        #region //寄送Mail
                        var mailConfig = new MailConfig
                        {
                            Host = item.Host,
                            Port = Convert.ToInt32(item.Port),
                            SendMode = Convert.ToInt32(item.SendMode),
                            From = item.MailFrom,
                            Subject = $"托外進貨週期性報表_{DateTime.Now:yyyy/MM/dd}",
                            Account = item.Account,
                            Password = item.Password,
                            MailTo = item.MailTo,
                            MailCc = item.MailCc,
                            MailBcc = item.MailBcc,
                            HtmlBody = $"附件為托外進貨週期性報表，日期:{DateTime.Now:yyyy/MM/dd}",
                            TextBody = "-",
                            FileInfo = new List<MailFile> { mailFile },
                            QcFileFlag = "N"
                        };
                        BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion

                    Response.Write(JObject.FromObject(new
                    {
                        status = "success",
                        msg = "報表已成功生成並寄送"
                    }).ToString());
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

        #region//ExcelRBPurchaseRequisitionProposal --粗胚建議請購报表 
        public void ExcelRBPurchaseRequisitionProposal(string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRBPurchaseRequisitionProposal(MtlItemNo, MtlItemName
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】粗胚建議請購报表 ";
                    string excelsheetName = "粗胚建議請購报表 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "產品品號", "產品品名", "產品規格", "ERP庫存單位", "第一分類", "第二分類", "第三分類", "第四分類"
                        , "最低補量", "補貨倍量", "安全存量", "庫存數量","建議採購數" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.InventoryUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.TypeOne.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.TypeTwo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.TypeThree.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.TypeFour.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.MinReorderQuantity.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ReorderMultiplier.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.SafetyStockQuantity.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.StockQuantity.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.RecommendedPurchaseQuantity.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 30;
                        //worksheet.Column(2).Width = 30;
                        //worksheet.Column(3).Width = 30;
                        //worksheet.Column(4).Width = 30;
                        //worksheet.Column(5).Width = 30;
                        //worksheet.Column(6).Width = 20;
                        //worksheet.Column(7).Width = 20;
                        //worksheet.Column(8).Width = 20;
                        //worksheet.Column(9).Width = 20;
                        //worksheet.Column(10).Width = 15;
                        //worksheet.Column(11).Width = 15;



                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelRBAnnualProcurementReport --年購粗胚明細 
        public void ExcelRBAnnualProcurementReport(string StartReqDate = "", string EndReqDate = "", string ReqName = ""
                    , string ReqNo = "", string ReqDepNo = "", string ReqUserNo = "", string SearchKey = ""
                    , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
                    , string PurUserNo = "", string PurVendorName = ""
                    , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetRBAnnualProcurementReport(StartReqDate, EndReqDate, ReqName
                    , ReqNo, ReqDepNo, ReqUserNo, SearchKey
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
                    , PurUserNo, PurVendorName
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】年購粗胚明細 ";
                    string excelsheetName = "年購粗胚明細 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "請購日期", "單據名稱", "請購單號", "確認碼", "確認者編號", "確認者", "部門編號", "請購部門"
                        , "人員工號", "請購人員", "請購備註", "品項序號", "品號" , "品名" , "規格" , "會計分類名稱" , "倉管分類名稱", "業務分類名稱", "生管分類名稱"
                        , "請購單位", "請購數量", "專案代號", "專案名稱", "品項備註", "請購單創建日", "請購單創建者編號","請購單創建者", "結案碼", "課稅別"
                        , "幣別", "匯率" , "營業稅率" , "人員帳號"  , "採購人員" , "廠商編號" , "廠商簡稱" , "採購數量" , "採購單位"
                        , "採購幣別", "採購單價", "採購金額", "本幣未稅採購金額"

                    };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ReqDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ReqName.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.ReqNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ConfirmationCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.ConfirmationUserNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ConfirmationUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ReqDepNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ReqDepName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ReqUserNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ReqUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.ReqRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.ReqLineNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.MtlItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.AccountingCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.WarehouseCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.BusinessCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.ManagementCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.ReqUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.ReqQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.ProjectCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.ProjectName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.ReqLineRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.ReqCreateDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.ReqCreateUserNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.ReqCreateUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.CaseCloseCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.TaxType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.CurrencyType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.ExchangeRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.BusinessTaxRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.PurUserNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.PurUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.PurVendorNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.PurVendorName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.PurQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.PurUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.PurCurrencyType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.PurPrice.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.PurAmount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.LocalCurrencyUntaxedPurAmount.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 30;
                        //worksheet.Column(2).Width = 30;
                        //worksheet.Column(3).Width = 30;
                        //worksheet.Column(4).Width = 30;
                        //worksheet.Column(5).Width = 30;
                        //worksheet.Column(6).Width = 20;
                        //worksheet.Column(7).Width = 20;
                        //worksheet.Column(8).Width = 20;
                        //worksheet.Column(9).Width = 20;
                        //worksheet.Column(10).Width = 15;
                        //worksheet.Column(11).Width = 15;



                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelDepProfitLossStatement --成本趋势分析报表 
        public void ExcelDepProfitLossStatement(string Department = "", string FiscalYear = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetDepProfitLossStatement(Department, FiscalYear
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】成本趋势分析报表 ";
                    string excelsheetName = "成本趋势分析报表 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "类别", "一月", "二月", "三月", "四月", "五月", "六月", "七月"
                        , "八月", "九月", "十月", "十一月","十二月" ,"总计" ,"占比"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;

                    using (var workbook = new XLWorkbook())
                    {

                        var worksheet = workbook.Worksheets.Add(excelsheetName);

                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //var imagePath = Server.MapPath("~/Content/images/logo.png");
                        //var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        //worksheet.Row(rowIndex).Height = 40;
                        //worksheet.Range(rowIndex, 1, rowIndex, 15).Merge().Style = titleStyle;
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion


                        #region //BODY
                        int totalRowIndex = 0;
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.classification.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.Jan.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.Feb.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.Mar.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Apr.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.May.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.Jun.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Jul.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Aug.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Sept.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Oct.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Nov.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.Dec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.Total.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.proportion.ToString();

                            // 判断是否为“总计”行，如果是则填充蓝色背景
                            if (item.classification == "总计")
                            {
                                totalRowIndex = rowIndex;
                                var customColor1 = XLColor.FromHtml("#D0DCDF");
                                for (int col = 1; col <= 15; col++)
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(col, rowIndex)).Style.Fill.BackgroundColor = customColor1;
                                }
                            }

                            if (item.classification == "变动占比" || item.classification == "变动%" || item.classification == "固定占比" || item.classification == "固定%")
                            {
                                var customColor2 = XLColor.FromHtml("#F3F3EE");
                                for (int col = 1; col <= 15; col++)
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(col, rowIndex)).Style.Fill.BackgroundColor = customColor2;
                                }
                            }

                        }
                        #endregion

                        #region //設定

                        #region //设置第 2 列到第 13 列的格式为数值的千位分隔符
                        for (int col = 2; col <= 14; col++)
                        {
                            worksheet.Column(col).Style.NumberFormat.Format = "#,##0.00";
                        }
                        #endregion

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();

                        #endregion


                        #region //凍結
                        //窗格、首欄、頂端列-·
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelCostAllocationTable --成本核算分配表 
        public void ExcelCostAllocationTable(string YearMonth = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetCostAllocationTable(YearMonth
                    , OrderBy, PageIndex, PageSize);
                #endregion

                logger.Info($"Query parameters - YearMonth: {YearMonth}, OrderBy: {OrderBy}, PageIndex: {PageIndex}, PageSize: {PageSize}");
                logger.Info($"Data request result: {dataRequest}");
                var dataObj = JObject.Parse(dataRequest)["data"];
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = YearMonth + "月" + "晶彩成本核算分配表 ";
                    string excelsheetName = "成本核算分配表 ";

                    dynamic[] LaborCostResult = JsonConvert.DeserializeObject<JObject[]>(dataObj["LaborCostResult"].ToString());
                    dynamic[] ManufacturCostResult = JsonConvert.DeserializeObject<JObject[]>(dataObj["ManufacturCostResult"].ToString());


                    string[] header = new string[] { "类别", "线别名称", "人工产能", "机器产能", "直接人工", "需摊人工", "全厂共摊人工", "分配到部门人工(一)"
                        , "分配到部门人工(二)", "人工成本", "有效制费成本", "无效制费成本","单位人工" ,"有效单位制费" ,"合计"};
                    string[] header1 = new string[] { "类别", "线别名称", "有效机时", "无效机时", "机器产能", "有效制费", "无效制费", "有效分配部门制费(一)"
                        , "无效分配部门制费(一)", "有效共摊制费", "无效共摊制费", "有效分配部门制费(二)","无效分配部门制费(二)" ,"全厂有效共摊制费","全厂无效共摊制费"
                         ,"有效分配部门制费(三)" ,"无效分配部门制费(三)" ,"有效制费合计","无效制费合计","制费成本合计","无效单位制费"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 15).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion


                        #region //BODY
                        int totalRowIndex = 0;
                        int startRowPurposeType1 = 0;  // 记录第 1 列（PurposeType）相同值的起始行
                        string currentPurposeType1 = "";   // 记录第 1 列（PurposeType）当前值
                        int startRowDistributedLabour = 0;  // 记录第 6 列（distributedLabour）相同值的起始行
                        string currentDistributedLabour = "";   // 记录第 6 列（distributedLabour）当前值
                        int startRowAllHandsShared = 0; // 记录第 7 列（allHandsShared）相同值的起始行
                        string currentAllHandsShared = "";  // 记录第 7 列（allHandsShared）当前值

                        foreach (var item in LaborCostResult)
                        {

                            startIndex = rowIndex + 1;
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.PurposeType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.ManualOutput.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MachineOutput.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.DirectLabor.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.distributedLabour.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.allHandsShared.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.departmentalLabor01.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.departmentalLabor02.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.LaborCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.EffectiveChargeCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.InvalidChargeCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.LaborPerUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.EffectiveUnitCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.TotalCost.ToString();

                            for (int col = 3; col <= 15; col++)
                            {
                                worksheet.Column(col).Style.NumberFormat.Format = "#,##0.00";
                            }

                            if (item.CategoryName == "合计")  // 判断是否为“合计”行，如果是则填充蓝色背景
                            {
                                totalRowIndex = rowIndex;
                                var customColor1 = XLColor.FromHtml("#D0DCDF");
                                for (int col = 1; col <= 15; col++)
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(col, rowIndex)).Style.Fill.BackgroundColor = customColor1;
                                }
                            }

                            if (currentPurposeType1 == "") // 处理第 1 列（PurposeType）的合并
                            {
                                currentPurposeType1 = item.PurposeType.ToString();
                                startRowPurposeType1 = rowIndex;
                            }
                            else if (currentPurposeType1 != item.PurposeType.ToString())
                            {
                                if (startRowPurposeType1 < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(1, startRowPurposeType1), BaseHelper.MergeNumberToChar(1, rowIndex - 1)).Merge();
                                }
                                currentPurposeType1 = item.PurposeType.ToString();
                                startRowPurposeType1 = rowIndex;
                            }

                            if (currentDistributedLabour == "") // 处理第 6 列（distributedLabour）的合并
                            {
                                currentDistributedLabour = item.distributedLabour.ToString();
                                startRowDistributedLabour = rowIndex;
                            }
                            else if (currentDistributedLabour != item.distributedLabour.ToString())
                            {
                                if (startRowDistributedLabour < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(6, startRowDistributedLabour), BaseHelper.MergeNumberToChar(6, rowIndex - 1)).Merge();
                                }
                                currentDistributedLabour = item.distributedLabour.ToString();
                                startRowDistributedLabour = rowIndex;
                            }

                            if (currentAllHandsShared == "")    // 处理第 7 列（allHandsShared）的合并
                            {
                                currentAllHandsShared = item.allHandsShared.ToString();
                                startRowAllHandsShared = rowIndex;
                            }
                            else if (currentAllHandsShared != item.allHandsShared.ToString())
                            {
                                if (startRowAllHandsShared < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(7, startRowAllHandsShared), BaseHelper.MergeNumberToChar(7, rowIndex - 1)).Merge();
                                }
                                currentAllHandsShared = item.allHandsShared.ToString();
                                startRowAllHandsShared = rowIndex;
                            }
                        }

                        if (startRowPurposeType1 < rowIndex - 1) // 处理最后一组相同值的合并（第 1 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(1, startRowPurposeType1), BaseHelper.MergeNumberToChar(1, rowIndex - 1)).Merge();
                        }

                        if (startRowDistributedLabour < rowIndex - 1) // 处理最后一组相同值的合并（第 6 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(6, startRowDistributedLabour), BaseHelper.MergeNumberToChar(6, rowIndex - 1)).Merge();
                        }

                        if (startRowAllHandsShared < rowIndex - 1)  // 处理最后一组相同值的合并（第 7 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(7, startRowAllHandsShared), BaseHelper.MergeNumberToChar(7, rowIndex - 1)).Merge();
                        }




                        rowIndex += 2;
                        for (int i = 0; i < header1.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header1[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }

                        int totalRowIndex1 = 0;
                        int startRowPurposeType = 0;  // 记录第 1 列（PurposeType）相同值的起始行
                        string currentPurposeType = "";   // 记录第 1 列（PurposeType）当前值
                        int startRowEffectiveCoayments = 0;  // 记录第 10 列（EffectiveCoayments）相同值的起始行
                        string currentEffectiveCoayments = "";   // 记录第 10 列（EffectiveCoayments）当前值
                        int startRowInvalidCoayments = 0; // 记录第 11 列（InvalidCoayments）相同值的起始行
                        string currentInvalidCoayments = "";  // 记录第 11 列（InvalidCoayments）当前值
                        int startRowEffectivePSE = 0;  // 记录第 14 列（EffectivePSE）相同值的起始行
                        string currentEffectivePSE = "";   // 记录第 14 列（EffectivePSE）当前值
                        int startRowInvalidPSE = 0; // 记录第 15 列（InvalidPSE）相同值的起始行
                        string currentInvalidPSE = "";  // 记录第 15 列（InvalidPSE）当前值

                        foreach (var item in ManufacturCostResult)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.PurposeType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.EffectiveTime.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.DeadTime.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MachineCapacity.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.EffectiveCharge.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.InvalidCharge.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.EffectiveChargeDep01.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.InvalidChargeDep01.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.EffectiveCoayments.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.InvalidCoayments.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.EffectiveChargeDep02.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.InvalidChargeDep02.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.EffectivePSE.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.InvalidPSE.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.EffectiveChargeDep03.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.InvalidChargeDep03.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.TotalEffectiveManufee.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.TotalInvalidManufee.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.TotalManufee.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.InvalidUnitfee.ToString();

                            for (int col = 3; col <= 21; col++)
                            {
                                worksheet.Column(col).Style.NumberFormat.Format = "#,##0.00";
                            }
                            // 判断是否为“合计”行，如果是则填充蓝色背景
                            if (item.CategoryName == "合计")
                            {
                                totalRowIndex1 = rowIndex;
                                var customColor1 = XLColor.FromHtml("#D0DCDF");
                                for (int col = 1; col <= 21; col++)
                                {
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(col, rowIndex)).Style.Fill.BackgroundColor = customColor1;
                                }
                            }

                            if (currentPurposeType == "") // 处理第 1 列（PurposeType）的合并
                            {
                                currentPurposeType = item.PurposeType.ToString();
                                startRowPurposeType = rowIndex;
                            }
                            else if (currentPurposeType != item.PurposeType.ToString())
                            {
                                if (startRowPurposeType < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(1, startRowPurposeType), BaseHelper.MergeNumberToChar(1, rowIndex - 1)).Merge();
                                }
                                currentPurposeType = item.PurposeType.ToString();
                                startRowPurposeType = rowIndex;
                            }

                            if (currentEffectiveCoayments == "") // 处理第 10 列（EffectiveCoayments）的合并
                            {
                                currentEffectiveCoayments = item.EffectiveCoayments.ToString();
                                startRowEffectiveCoayments = rowIndex;
                            }
                            else if (currentEffectiveCoayments != item.EffectiveCoayments.ToString())
                            {
                                if (startRowEffectiveCoayments < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(10, startRowEffectiveCoayments), BaseHelper.MergeNumberToChar(10, rowIndex - 1)).Merge();
                                }
                                currentEffectiveCoayments = item.EffectiveCoayments.ToString();
                                startRowEffectiveCoayments = rowIndex;
                            }

                            if (currentInvalidCoayments == "")    // 处理第 11 列（InvalidCoayments）的合并
                            {
                                currentInvalidCoayments = item.InvalidCoayments.ToString();
                                startRowInvalidCoayments = rowIndex;
                            }
                            else if (currentInvalidCoayments != item.InvalidCoayments.ToString())
                            {
                                if (startRowInvalidCoayments < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(11, startRowInvalidCoayments), BaseHelper.MergeNumberToChar(11, rowIndex - 1)).Merge();
                                }
                                currentInvalidCoayments = item.InvalidCoayments.ToString();
                                startRowInvalidCoayments = rowIndex;
                            }

                            if (currentEffectivePSE == "") // 处理第 14 列（EffectivePSE）的合并
                            {
                                currentEffectivePSE = item.EffectivePSE.ToString();
                                startRowEffectivePSE = rowIndex;
                            }
                            else if (currentEffectivePSE != item.EffectivePSE.ToString())
                            {
                                if (startRowEffectivePSE < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(14, startRowEffectivePSE), BaseHelper.MergeNumberToChar(14, rowIndex - 1)).Merge();
                                }
                                currentEffectivePSE = item.EffectivePSE.ToString();
                                startRowEffectivePSE = rowIndex;
                            }

                            if (currentInvalidPSE == "")    // 处理第 15 列（InvalidPSE）的合并
                            {
                                currentInvalidPSE = item.InvalidPSE.ToString();
                                startRowInvalidPSE = rowIndex;
                            }
                            else if (currentInvalidPSE != item.InvalidPSE.ToString())
                            {
                                if (startRowInvalidPSE < rowIndex - 1)
                                {
                                    worksheet.Range(BaseHelper.MergeNumberToChar(15, startRowInvalidPSE), BaseHelper.MergeNumberToChar(15, rowIndex - 1)).Merge();
                                }
                                currentInvalidPSE = item.InvalidPSE.ToString();
                                startRowInvalidPSE = rowIndex;
                            }
                        }

                        if (startRowPurposeType < rowIndex - 1) // 处理最后一组相同值的合并（第 1 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(1, startRowPurposeType), BaseHelper.MergeNumberToChar(1, rowIndex - 1)).Merge();
                        }

                        if (startRowEffectiveCoayments < rowIndex - 1) // 处理最后一组相同值的合并（第 10 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(10, startRowEffectiveCoayments), BaseHelper.MergeNumberToChar(10, rowIndex - 1)).Merge();
                        }

                        if (startRowInvalidCoayments < rowIndex - 1)  // 处理最后一组相同值的合并（第 11 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(11, startRowInvalidCoayments), BaseHelper.MergeNumberToChar(11, rowIndex - 1)).Merge();
                        }
                        if (startRowEffectivePSE < rowIndex - 1) // 处理最后一组相同值的合并（第 14 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(14, startRowEffectivePSE), BaseHelper.MergeNumberToChar(14, rowIndex - 1)).Merge();
                        }

                        if (startRowInvalidPSE < rowIndex - 1)  // 处理最后一组相同值的合并（第 15 列）
                        {
                            worksheet.Range(BaseHelper.MergeNumberToChar(15, startRowInvalidPSE), BaseHelper.MergeNumberToChar(15, rowIndex - 1)).Merge();
                        }
                        #endregion

                        #region //設定

                        #region //设置第 2 列到第 13 列的格式为数值的千位分隔符
                        // for (int col = 3; col <= 15; col++)
                        // {
                        //     worksheet.Column(col).Style.NumberFormat.Format = "#,##0";
                        // }
                        #endregion

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();

                        #endregion


                        #region //凍結
                        //窗格、首欄、頂端列
                        worksheet.SheetView.FreezeRows(2);
                        worksheet.SheetView.FreezeColumns(2);
                        #endregion

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

        #region//ExcelSalesCostDetails --晶彩銷售成本明細 
        public void ExcelSalesCostDetails(string RoErpPrefix = "", string RoErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSalesCostDetails(RoErpPrefix, RoErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】晶彩銷售成本明細 ";
                    string excelsheetName = "晶彩銷售成本明細 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "單據日期", "銷貨日期", "單據名稱", "單號", "單頭備註", "確認碼", "確認者編號", "確認者"
                        , "結案碼", "客戶編號", "客戶簡稱", "部門編號", "負責部門" , "業務編號" , "負責業務" , "品項序號" , "品號", "品名", "規格"
                        , "會計分類編號", "會計分類名稱" , "倉管分類編號" , "倉管分類名稱"  , "業務分類編號" , "業務分類名稱" , "生管分類編號" , "生管分類名稱" , "庫別"
                        , "課稅別", "幣別", "匯率", "營業稅率" ,"數量", "單位", "單價", "金額" ,"單位成本", "材料成本", "人工成本", "製費成本"
                        , "加工成本", "本幣未稅金額" ,"銷貨成本", "本幣未稅毛利", "品項備註"

                    };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //var imagePath = Server.MapPath("~/Content/images/logo.png");
                        //var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        //worksheet.Row(rowIndex).Height = 40;
                        //worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DocDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.SalesDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DocName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.DocNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.HeaderRemarks.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ConfirmationCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ConfirmerID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ConfirmerName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.CaseClosureCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.CustomerNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.CustomerAbbreviation.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.DepID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.DepName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.BusinessID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.BusinessName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.LineNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.ItemID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.ItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.ItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.AccountingCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.AccountingCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.WarehouseCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.WarehouseCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.BusinessCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.BusinessCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.ManagementCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.ManagementCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.WarehouseName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.TaxType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.CurrencyType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.ExRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.BusinessTaxRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.Qty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.Unite.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.Price.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.Amount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.UnitCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.MaterialCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.LaborCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.ManufacturingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.ProcessingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.UntaxedLocalAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.SalesCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.LocalUntaxedGP.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.ItemRemarks.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

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

        #region//ExcelShipmentCostDetails --晶彩出貨成本明細 
        public void ExcelShipmentCostDetails(string TsnErpPrefix = "", string TsnErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetShipmentCostDetails(TsnErpPrefix, TsnErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】晶彩出貨成本明細 ";
                    string excelsheetName = "晶彩出貨成本明細 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "單據日期", "異動日期", "單據名稱", "單號", "單頭備註", "確認碼", "確認者編號", "確認者"
                        , "結案碼", "對象","對象編號", "對象簡稱", "部門編號", "負責部門" , "人員編號" , "負責人員" , "品項序號" , "品號", "品名", "規格"
                        , "會計分類編號", "會計分類名稱" , "倉管分類編號" , "倉管分類名稱"  , "業務分類編號" , "業務分類名稱" , "生管分類編號" , "生管分類名稱" , "轉出庫", "轉入庫"
                        , "課稅別", "幣別", "匯率", "營業稅率" ,"數量", "單位", "單價", "金額" ,"單位成本", "材料成本", "人工成本", "製費成本"
                        , "加工成本", "暫出未稅金額" ,"銷貨成本", "暫出未稅毛利", "品項備註"

                    };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //var imagePath = Server.MapPath("~/Content/images/logo.png");
                        //var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        //worksheet.Row(rowIndex).Height = 40;
                        //worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            //var ChangeDate = item.ChangeDate != null ? item.ChangeDate.ToString() : "";

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DocDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ChangeDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DocName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.DocNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.HeaderRemarks.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ConfirmationCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ConfirmerID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ConfirmerName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.CaseClosureCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ObjectName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.ObjectID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.ObjectAbbr.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.DepID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.DepName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.EmployeeID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.EmployeeName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.LineNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.ItemID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.ItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.ItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.AccountingCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.AccountingCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.WarehouseCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.WarehouseCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.BusinessCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.BusinessCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.ManagementCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.ManagementCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.WarehouseTransferOut.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.WarehouseTransferIn.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.TaxType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.CurrencyType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.ExRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.BusinessTaxRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.Qty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.Unite.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.Price.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.Amount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.UnitCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.MaterialCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.LaborCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.ManufacturingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.ProcessingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.TempOutExTaxAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.SalesCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.TempOutExTaxGP.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.ItemRemarks.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

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

        #region//ExcelReimbursementCostDetails --晶彩歸還成本明細 
        public void ExcelReimbursementCostDetails(string TsrnErpPrefix = "", string TsrnErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetReimbursementCostDetails(TsrnErpPrefix, TsrnErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】晶彩歸還成本明細 ";
                    string excelsheetName = "晶彩歸還成本明細 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "單據日期", "單據名稱", "單號", "單頭備註", "確認碼", "確認者編號", "確認者"
                        , "對象","對象編號", "對象簡稱", "部門編號", "負責部門" , "人員編號" , "負責人員" , "品項序號" , "品號", "品名", "規格"
                        , "會計分類編號", "會計分類名稱" , "倉管分類編號" , "倉管分類名稱"  , "業務分類編號" , "業務分類名稱" , "生管分類編號" , "生管分類名稱" , "轉入庫", "轉出庫"
                        , "課稅別", "幣別", "匯率", "營業稅率" ,"類型","贈/備品量","贈/備品包裝量","計價數量","計價單位"
                        , "數量", "單位", "單價", "金額" ,"單位成本", "材料成本", "人工成本", "製費成本"
                        , "加工成本", "未稅金額" ,"銷貨成本", "未稅毛利", "品項備註"

                    };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //var imagePath = Server.MapPath("~/Content/images/logo.png");
                        //var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        //worksheet.Row(rowIndex).Height = 40;
                        //worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            //var ChangeDate = item.ChangeDate != null ? item.ChangeDate.ToString() : "";

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DocDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DocName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DocNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.HeaderRemarks.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.ConfirmationCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ConfirmerID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ConfirmerName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ObjectName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ObjectID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ObjectAbbr.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.DepID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.DepName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.EmployeeID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.EmployeeName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.LineNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.ItemID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.ItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.ItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.AccountingCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.AccountingCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.WarehouseCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.WarehouseCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.BusinessCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.BusinessCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.ManagementCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.ManagementCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.WarehouseTransferIn.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.WarehouseTransferOut.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.TaxType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.CurrencyType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.ExRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.BusinessTaxRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.GSType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.GSQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.GSPackQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.BillingQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.PricingUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.Qty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.Unite.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.Price.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.Amount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.UnitCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.MaterialCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.LaborCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.ManufacturingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.ProcessingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.UntaxedAmount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(48, rowIndex)).Value = item.SalesCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(49, rowIndex)).Value = item.UntaxedGP.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(50, rowIndex)).Value = item.ItemRemarks.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

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

        #region//ExcelSalesReturnsAllowancesCostDetails --晶彩銷退/折讓成本明細 
        public void ExcelSalesReturnsAllowancesCostDetails(string RtErpPrefix = "", string RtErpNo = "", string CustomerNo = ""
            , string SalesmenNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetSalesReturnsAllowancesCostDetails(RtErpPrefix, RtErpNo, CustomerNo
                    , SalesmenNo, MtlItemNo, StartDate, EndDate
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】晶彩銷退/折讓成本明細 ";
                    string excelsheetName = "晶彩銷退折讓成本明細 ";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    //List<MoProgressDetailResult> detailReportDatas = JsonConvert.DeserializeObject<List<MoProgressDetailResult>>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "單據日期", "銷退日", "類型","單據名稱","單號", "單頭備註", "確認碼", "確認者編號", "確認者", "結帳碼"
                        , "客戶編碼", "客戶簡稱", "部門編號", "負責部門" , "業務編號" , "負責業務" , "品項序號" , "品號", "品名", "規格"
                        , "會計分類編號", "會計分類名稱" , "倉管分類編號" , "倉管分類名稱"  , "業務分類編號" , "業務分類名稱" , "生管分類編號" , "生管分類名稱" , "退入庫別"
                        , "數量類型","贈/備品量","贈/備品包裝量","計價數量","計價單位"
                        , "數量", "單位", "單價", "金額" ,"單位成本", "材料成本", "人工成本", "製費成本"
                        , "加工成本", "本幣未稅金額" ,"銷貨成本", "本幣未稅毛利", "品項備註"
                        , "單據創建日", "單據創建者編號" ,"單據創建者", "單據修改日", "單據修改者編號", "單據修改者"

                    };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //var imagePath = Server.MapPath("~/Content/images/logo.png");
                        //var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        //worksheet.Row(rowIndex).Height = 40;
                        //worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;

                            //var ChangeDate = item.ChangeDate != null ? item.ChangeDate.ToString() : "";

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DocDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.SRDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.SRAType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.DocName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.DocNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.HeaderRemarks.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ConfirmationCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ConfirmerID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ConfirmerName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.CaseClosureCode.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.CustomerID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.CustomerAbbr.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.DepID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.DepName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.BusinessID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.BusinessName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.LineNum.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.ItemID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.ItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.ItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.AccountingCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.AccountingCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.WarehouseCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.WarehouseCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.BusinessCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.BusinessCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.ManagementCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.ManagementCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.RIWarehouse.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.GSType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.GSQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.GSPackQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.BillingQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.PricingUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.Qty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.Unite.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.Price.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.Amount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.UnitCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.MaterialCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.LaborCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.ManufacturingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.ProcessingCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.UntaxedLocalAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.SalesCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.LocalUntaxedGP.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.ItemRemarks.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(48, rowIndex)).Value = item.DocCreatDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(49, rowIndex)).Value = item.DocCreatUserID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(50, rowIndex)).Value = item.DocCreatUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(51, rowIndex)).Value = item.DocUpDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(52, rowIndex)).Value = item.DocUpUserID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(53, rowIndex)).Value = item.DocUpUserName.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(1);
                        #endregion

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

        #region //ExcelOspDetailSummary 托外單交期追蹤輸出Excel
        public void ExcelOspDetailSummary(string OspNo = "", string SupplierNo = "", string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = ""
            , string OspCreatStartDate = "", string OspCreatEndDate = "", string Status = "", string DelayStatus = "", int UserId = -1, int DepartmentId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OspDetailSummary", "read,excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetOspDetailSummary(OspNo, SupplierNo, WoErpFullNo, MtlItemNo, MtlItemName
                    , OspCreatStartDate, OspCreatEndDate, Status, DelayStatus, UserId, DepartmentId
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】托外單交期追蹤";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "供應商", "託外生產單號", "交期狀況", "單據日", "預計回廠日"
                        , "最後入庫日", "託外單數量", "託外入庫數量", "單據創建者", "品號", "品名", "規格", "製令", "製程"
                         };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;


                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            rowIndex++;

                            int outUse = rowIndex;

                            string json = item["DetailsJson"]?.ToString() ?? "[]"; // 確保取出的是字串形式的 JSON 陣列
                            JArray jsonArray = JArray.Parse(json);
                            dynamic[] dataDetail = jsonArray.ToObject<dynamic[]>();

                            foreach (var item2 in dataDetail)
                            {

                                string mtlItemNo = item2.MtlItemNo.ToString();
                                string mtlItemName = item2.MtlItemName.ToString();
                                string mtlItemSpec = item2.MtlItemSpec.ToString();
                                string woErpFull = item2.WoErpFull.ToString();
                                string processCodeNames = item2.ProcessCodeNames.ToString();

                                var resultList1 = new[]
                                {
                                    new {
                                        mtlItemNo,
                                        mtlItemName,
                                        mtlItemSpec,
                                        woErpFull,
                                        processCodeNames
                                    }
                                };

                                var item3 = resultList1[0]; // 因為你目前陣列只有一筆

                                var properties1 = item3.GetType().GetProperties(); // 用 reflection 取得屬性
                                for (int i = 0; i < properties1.Length; i++)
                                {
                                    var value = properties1[i].GetValue(item3)?.ToString() ?? "";
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(i + 10, rowIndex)).Value = value;

                                }
                                rowIndex++;

                            }
                            rowIndex--;

                            string OrderStatus = "";
                            if (item.CloseStatus == "N")
                            {
                                OrderStatus = "未結案";
                            }
                            else
                            {
                                OrderStatus = "結案";
                            }
                            string supperlier = item.Supperlier.ToString();
                            string ospNo = item.OspNo.ToString();
                            string ospDate = item.OspDate.ToString();
                            string delayStatus = item.DelayStatus.ToString();
                            string expectedDate = item.ExpectedDate.ToString();
                            string receiptDate = item.ReceiptDate.ToString();
                            string qty = item.TrueQty.ToString();
                            string receiptQty = item.TrueReceiptQty.ToString();
                            string createUser = item.CreateUser.ToString();

                            var resultList = new[]
                            {
                                new {
                                    supperlier,
                                    ospNo,
                                    delayStatus,
                                    ospDate,
                                    expectedDate,
                                    receiptDate,
                                    qty,
                                    receiptQty,
                                    createUser
                                }
                            };


                            var item1 = resultList[0]; // 因為你目前陣列只有一筆

                            var properties = item1.GetType().GetProperties(); // 用 reflection 取得屬性
                            for (int i = 0; i < properties.Length; i++)
                            {
                                var value = properties[i].GetValue(item1)?.ToString() ?? "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, outUse)).Value = value;
                            }

                            worksheet.Range($"A{outUse}:A{rowIndex}").Merge();
                            worksheet.Range($"B{outUse}:B{rowIndex}").Merge();
                            worksheet.Range($"C{outUse}:C{rowIndex}").Merge();
                            worksheet.Range($"D{outUse}:D{rowIndex}").Merge();
                            worksheet.Range($"E{outUse}:E{rowIndex}").Merge();
                            worksheet.Range($"F{outUse}:F{rowIndex}").Merge();
                            worksheet.Range($"G{outUse}:G{rowIndex}").Merge();
                            worksheet.Range($"H{outUse}:H{rowIndex}").Merge();
                            worksheet.Range($"I{outUse}:I{rowIndex}").Merge();
                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

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
                        worksheet.SheetView.FreezeRows(1);
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

        #region //匯出過站移轉還原數據 ExcelStationTransferRestoreData
        public void ExcelTransferDetailData(int ModeId = -1, string WoErpFullNo = "", string MtlItemName = "", string StartDate = "", string EndDate = "",
            int UserId = -1, string BarcodeNo = "", string ProcessAlias = "", string ShopDesc = "", string MachineDesc = "", string OrderBy = "")
        {
            try
            {
                #region //權限檢查
               // WebApiLoginCheck("StationTransferRestore", "excel");
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetTransferDetailData(ModeId, WoErpFullNo, MtlItemName, StartDate, EndDate,
                    UserId, BarcodeNo, ProcessAlias, ShopDesc, MachineDesc, OrderBy, -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】過站移轉還原數據";
                    string excelsheetName = "過站移轉還原數據";

                    List<dynamic> detailReportDatas = ((JArray)JObject.Parse(dataRequest)["data"]).ToObject<List<dynamic>>();

                    string[] header = new string[] {
                "模式名稱", "製令", "計畫數", "製程別名", "部番", "條碼", 
                "原過站數量", "NG數", "開工人員", "開工時間", "完工人員", "完工時間",
                "加工工時", "機台", "車間", "平均工時/P", "班別", "移出數量", "移入數量", "還原過站數量", "完工年", "完工月", "完工日"
            };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in detailReportDatas)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ModeName.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.WoErpFullNo.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.PlanQty.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ProcessAlias.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MtlItemName.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.BarcodeNo.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.StatusName.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.StationQty.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.NgQty.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.StartUser.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.StartDate.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.FinishUser.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.FinishDate.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.CycleTime.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.MachineDesc.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.ShopDesc.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.AvgCycleTimePerPiece.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.ShiftName.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.TransferOutQty.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.TransferInQty.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.OriginStationQty.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.FinishYear.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.FinishMonth.ToString(); ;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.FinishDay.ToString(); ;
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//ExcelBarcodeBindingData 條碼綁定資訊匯出       
        public void ExcelBarcodeBindingData(string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = "", string BatchBarcodeNo = "", string LensBarcodeNo = "", string ProcessAlias = "", string StartDate = "", string EndDate = "", string UserNo = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeBindingData(MtlItemNo, MtlItemName, WoErpFullNo, BatchBarcodeNo, LensBarcodeNo, ProcessAlias, StartDate, EndDate, UserNo, OrderBy, PageIndex, PageSize);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】條碼綁定資訊";
                    string excelsheetName = "條碼綁定資訊";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "品號", "品名", "工單編號", "批量條碼", "鏡頭條碼", "製程名稱", "綁定時間", "批量條碼數量", "人員" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 9).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MtlItemName.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.WoErpFullNo.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.BatchBarcodeNo.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.LensBarcodeNo.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ProcessAlias.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.CreateDate.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.BatchBarcodeQty.ToString(); 
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.UserName.ToString(); 
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        //worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region//LCMObsoleteStockProvisionDetailExcel JC實時庫存LCM呆滯提列

        public void LCMObsoleteStockProvisionDetailExcel(string MtlItemNo = ""
            , string AccountingCategoryNo = "", string WarehouseCategoryNo = "", string BusinessCategoryNo = "", string ManagementCategoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        {
            try
            {
                //WebApiLoginCheck("ToolTransactions", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetLCMObsoleteStockProvisionDetail(MtlItemNo
                    , AccountingCategoryNo, WarehouseCategoryNo, BusinessCategoryNo, ManagementCategoryNo
                    , OrderBy, PageIndex, PageSize, RedirectType);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】JC實時庫存LCM呆滯提列";
                    string excelsheetName = "JC實時庫存LCM呆滯提列";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["result"].ToString());
                    string[] header = new string[] { "品號", "品名", "規格", "庫存單位", "庫存數量", "庫存金額", "單位成本", "庫別編號", "庫別名稱", "儲存位置", "庫齡條件", "呆滯提列金額"
                    , "存貨淨額", "存貨淨單價", "課稅別", "幣別", "匯率", "營業稅率", "最近訂單單價", "最近訂單本幣未稅單價", "淨變現單價", "單價差額", "差額金額"
                    , "判斷", "會計分類編號", "會計分類名稱", "倉管分類編號", "倉管分類名稱", "業務分類編號", "業務分類名稱", "生管分類編號", "生管分類名稱", "品號創建日", "品號創建者編號", "品號創建者名"
                    , "品號修改日", "品號修改者編號", "品號修改者名", "最近入庫日", "最近出庫日", "進出創建日", "進出創建者編號", "進出創建者名", "進出修改日", "進出修改者編號", "進出修改者名"
                    , "品號生效日", "品號失效日"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        //var imagePath = Server.MapPath("~/Content/images/logo.png");
                        //var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 5)).Scale(0.07);
                        //worksheet.Row(rowIndex).Height = 40;
                        //worksheet.Range(rowIndex, 1, rowIndex, 13).Merge().Style = titleStyle;
                        //rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;

                        foreach (var item in data)
                        {
                            rowIndex++;

                            startIndex = rowIndex;

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ItemID.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.ItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.InventoryUnit.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.InventoryQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.InventoryAmount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.UnitCost.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.InventoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.InventoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.StorageLocation.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.InventoryAgeCondition.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.StagnantProvisionAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.InventoryNetAmount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.InventoryNetPrice.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.TaxType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.CurrencyType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.ExRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.BusinessTaxRate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.RecentOrderPrice.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.RecentOrdLocCurExTaxPrice.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.NetRealizablePrice.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.UnitPriceVariance.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.VarianceAmount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.Decision.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.AccountingCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.AccountingCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.WarehouseCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.WarehouseCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.BusinessCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.BusinessCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.ManagementCategoryNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.ManagementCategoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.ProductCreatDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.ProductCreatNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.ProductCreatName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.ProductUpdateDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.ProductUpdateNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.ProductUpdateName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.ReWarInDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.ReWarOutDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.ReWarCreatDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.ReWarCreatNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.ReWarCreatName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.ReWarUpdateDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.ReWarUpdateNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.ReWarUpdateName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.ProductEffectiveDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(48, rowIndex)).Value = item.ProductExpiryDate.ToString();




                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                            worksheet.Range(startIndex, 5, rowIndex, 5).Merge();
                            worksheet.Range(startIndex, 6, rowIndex, 6).Merge();
                            worksheet.Range(startIndex, 7, rowIndex, 7).Merge();
                            worksheet.Range(startIndex, 8, rowIndex, 8).Merge();
                            worksheet.Range(startIndex, 9, rowIndex, 9).Merge();
                            worksheet.Range(startIndex, 10, rowIndex, 10).Merge();
                            worksheet.Range(startIndex, 11, rowIndex, 11).Merge();
                            worksheet.Range(startIndex, 12, rowIndex, 12).Merge();
                            worksheet.Range(startIndex, 13, rowIndex, 13).Merge();
                            worksheet.Range(startIndex, 14, rowIndex, 14).Merge();
                            worksheet.Range(startIndex, 15, rowIndex, 15).Merge();
                            worksheet.Range(startIndex, 16, rowIndex, 16).Merge();
                            worksheet.Range(startIndex, 17, rowIndex, 17).Merge();
                            worksheet.Range(startIndex, 18, rowIndex, 18).Merge();
                            worksheet.Range(startIndex, 19, rowIndex, 19).Merge();
                            worksheet.Range(startIndex, 20, rowIndex, 20).Merge();
                            worksheet.Range(startIndex, 21, rowIndex, 21).Merge();
                            worksheet.Range(startIndex, 22, rowIndex, 22).Merge();
                            worksheet.Range(startIndex, 23, rowIndex, 23).Merge();
                            worksheet.Range(startIndex, 24, rowIndex, 24).Merge();
                            worksheet.Range(startIndex, 25, rowIndex, 25).Merge();
                            worksheet.Range(startIndex, 26, rowIndex, 26).Merge();
                            worksheet.Range(startIndex, 27, rowIndex, 27).Merge();
                            worksheet.Range(startIndex, 28, rowIndex, 28).Merge();
                            worksheet.Range(startIndex, 29, rowIndex, 29).Merge();
                            worksheet.Range(startIndex, 30, rowIndex, 30).Merge();
                            worksheet.Range(startIndex, 31, rowIndex, 31).Merge();
                            worksheet.Range(startIndex, 32, rowIndex, 32).Merge();
                            worksheet.Range(startIndex, 33, rowIndex, 33).Merge();
                            worksheet.Range(startIndex, 34, rowIndex, 34).Merge();
                            worksheet.Range(startIndex, 35, rowIndex, 35).Merge();
                            worksheet.Range(startIndex, 36, rowIndex, 36).Merge();
                            worksheet.Range(startIndex, 37, rowIndex, 37).Merge();
                            worksheet.Range(startIndex, 38, rowIndex, 38).Merge();
                            worksheet.Range(startIndex, 39, rowIndex, 39).Merge();
                            worksheet.Range(startIndex, 40, rowIndex, 40).Merge();
                            worksheet.Range(startIndex, 41, rowIndex, 41).Merge();
                            worksheet.Range(startIndex, 42, rowIndex, 42).Merge();
                            worksheet.Range(startIndex, 43, rowIndex, 43).Merge();
                            worksheet.Range(startIndex, 44, rowIndex, 44).Merge();
                            worksheet.Range(startIndex, 45, rowIndex, 45).Merge();
                            worksheet.Range(startIndex, 46, rowIndex, 46).Merge();
                            worksheet.Range(startIndex, 47, rowIndex, 47).Merge();
                            worksheet.Range(startIndex, 48, rowIndex, 48).Merge();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;
                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //ExcelProductCostReport 產品成本分析表
        public void ExcelProductCostReport(string MtlItemNo = "", string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("OspDetailSummary", "read,excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductCostReport(MtlItemNo, StardDay, EndDay);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】產品成本分析表";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "產品品號", "品名", "規格", "單位", "生產入庫"
                        ,"託外進貨" ,"材料在製約量" ,"人工製費在製約量" ,"加工費用在製約量"
                        ,"材料成本" ,"單位材料成本" ,"材料成本(％)" ,"人工成本" ,"單位人工成本"
                        ,"人工成本(％)" ,"製造費用" ,"單位製造費用" ,"製造費用(％)" ,"生產成本"
                        ,"單位生產成本" ,"本階人工成本" ,"本階製造費用" ,"本階加工費用"
                         };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion
                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;


                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            var rowData = new[]
                            {
                                item.MtlItemNo?.ToString() ?? "",
                                item.MtlItemName?.ToString() ?? "",
                                item.MtlItemSpec?.ToString() ?? "",
                                item.UomNo?.ToString() ?? "",
                                item.ProdInQty?.ToString() ?? "",
                                item.OutsourceInQty?.ToString() ?? "",
                                item.MaterialWIPReserve?.ToString() ?? "",
                                item.LaborWIPReserve?.ToString() ?? "",
                                item.MfgCostWIPReserve?.ToString() ?? "",
                                item.MaterialCost?.ToString() ?? "",
                                item.UnitMaterialCost?.ToString() ?? "",
                                item.MaterialCostRate?.ToString() ?? "",
                                item.LaborCost?.ToString() ?? "",
                                item.UnitLaborCost?.ToString() ?? "",
                                item.LaborCostRate?.ToString() ?? "",
                                item.MfgOverhead?.ToString() ?? "",
                                item.UnitMfgCost?.ToString() ?? "",
                                item.MfgCostRate?.ToString() ?? "",
                                item.TotalProdCost?.ToString() ?? "",
                                item.UnitProdCost?.ToString() ?? "",
                                item.CurrStageLaborCost?.ToString() ?? "",
                                item.CurrStageMfgOverhead?.ToString() ?? ""
                            };

                            rowIndex++;

                            for (int i = 0; i < rowData.Length; i++)
                            {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, rowIndex)).Value = rowData[i];
                            }
                        }

                        //foreach (var item in data)
                        //{
                        //    string ErpMtlItemNo = item.MtlItemNo.ToString();
                        //    string MtlItemName = item.MtlItemName.ToString();
                        //    string MtlItemSpec = item.MtlItemSpec.ToString();
                        //    string UomNo = item.UomNo.ToString();
                        //    string ProdInQty = item.ProdInQty.ToString();
                        //    string OutsourceInQty = item.OutsourceInQty.ToString();
                        //    string MaterialWIPReserve = item.MaterialWIPReserve.ToString();
                        //    string LaborWIPReserve = item.LaborWIPReserve.ToString();
                        //    string MfgCostWIPReserve = item.MfgCostWIPReserve.ToString();
                        //    string MaterialCost = item.MaterialCost.ToString();
                        //    string UnitMaterialCost = item.UnitMaterialCost.ToString();
                        //    string MaterialCostRate = item.MaterialCostRate.ToString();
                        //    string LaborCost = item.LaborCost.ToString();
                        //    string UnitLaborCost = item.UnitLaborCost.ToString();
                        //    string LaborCostRate = item.LaborCostRate.ToString();
                        //    string MfgOverhead = item.MfgOverhead.ToString();
                        //    string UnitMfgCost = item.UnitMfgCost.ToString();
                        //    string MfgCostRate = item.MfgCostRate.ToString();
                        //    string TotalProdCost = item.TotalProdCost.ToString();
                        //    string UnitProdCost = item.UnitProdCost.ToString();
                        //    string CurrStageLaborCost = item.CurrStageLaborCost.ToString();
                        //    string CurrStageMfgOverhead = item.CurrStageMfgOverhead.ToString();


                        //    var resultList = new[]
                        //    {
                        //        new {
                        //            ErpMtlItemNo,
                        //            MtlItemName,
                        //            MtlItemSpec,
                        //            UomNo,
                        //            ProdInQty,
                        //            OutsourceInQty,
                        //            MaterialWIPReserve,
                        //            LaborWIPReserve,
                        //            MfgCostWIPReserve,
                        //            MaterialCost,
                        //            UnitMaterialCost,
                        //            MaterialCostRate,
                        //            LaborCost,
                        //            UnitLaborCost,
                        //            LaborCostRate,
                        //            MfgOverhead,
                        //            UnitMfgCost,
                        //            MfgCostRate,
                        //            TotalProdCost,
                        //            UnitProdCost,
                        //            CurrStageLaborCost,
                        //            CurrStageMfgOverhead
                        //        }
                        //    };

                        //    rowIndex++;

                        //    var item1 = resultList[0]; // 因為你目前陣列只有一筆

                        //    var properties = item1.GetType().GetProperties(); // 用 reflection 取得屬性
                        //    for (int i = 0; i < properties.Length; i++)
                        //    {
                        //        var value = properties[i].GetValue(item1)?.ToString() ?? "";
                        //        worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, rowIndex)).Value = value;
                        //    }
                        //}

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

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
                        worksheet.SheetView.FreezeRows(1);
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

        #region //ExcelProductReceivingCost 產品入庫成本分析表
        public void ExcelProductReceivingCost(string MtlItemNo = "", string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("OspDetailSummary", "read,excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductReceivingCost(MtlItemNo, StardDay, EndDay);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】產品入庫成本分析表";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "產品品號", "品名", "規格", "單位", "製令編號", "生產入庫"
                        ,"託外進貨","材料成本","人工成本","製造費用","加工費用","生產成本" ,"單位生產成本" 
                         };

                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;


                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            var rowData = new[]
                            {
                                item.MtlItemNo?.ToString() ?? "",
                                item.MtlItemName?.ToString() ?? "",
                                item.MtlItemSpec?.ToString() ?? "",
                                item.UomNo?.ToString() ?? "",
                                item.ErpFUllNo?.ToString() ?? "",
                                item.ProdInQty?.ToString() ?? "",
                                item.OutsourceInQty?.ToString() ?? "",
                                item.MaterialCost?.ToString() ?? "",
                                item.LaborCost?.ToString() ?? "",
                                item.MfgOverhead?.ToString() ?? "",
                                item.ProcessingFee?.ToString() ?? "",
                                item.TotalProdCost?.ToString() ?? "",
                                item.UnitProdCost?.ToString() ?? "",
                            };

                            rowIndex++;

                            for (int i = 0; i < rowData.Length; i++)
                            {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, rowIndex)).Value = rowData[i];
                            }
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

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
                        worksheet.SheetView.FreezeRows(1);
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

        #region //ExcelPoStatusReport 採購進貨狀況分析表
        public void ExcelPoStatusReport(string PoFullNo = "", string MtlItemNo = "", string SupplierNo = "", string CloseCode = "", string PoStatus = "", string DeliveryStatus = "", string StardDay = "", string EndDay = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        { 
            try
            {
                WebApiLoginCheck("ExcelExportManagement", "read,excel");
                List<string> OsrIdList = new List<string>();
                //增設 "供應商","單據日" 兩個欄位,前端動態報表要增加
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetPoStatusReport(PoFullNo,MtlItemNo, SupplierNo, CloseCode, PoStatus, DeliveryStatus, StardDay, EndDay, OrderBy, PageIndex, PageSize, RedirectType);
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】採購進貨狀況分析表";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["result"].ToString());
                    string[] header = new string[] { "採購單號", "品號", "品名", "規格", "交期狀況", "入庫狀況", "採購交期"
                        ,"最近進貨日","採購數量","進貨數量","驗收數","驗退數","單據日","供應商","結案狀態"
                         };

                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;


                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            var rowData = new[]
                            {
                                item.FullNo?.ToString() ?? "",
                                item.MtlItemNo?.ToString() ?? "",
                                item.MtlItemName?.ToString() ?? "",
                                item.MtlItemSpec?.ToString() ?? "",
                                item.DeliveryStatus?.ToString() ?? "",
                                item.PoStatus?.ToString() ?? "",
                                item.DeliveryDate?.ToString() ?? "",
                                item.InDay?.ToString() ?? "",
                                item.PuQty?.ToString() ?? "",
                                item.InTotalQty?.ToString() ?? "",
                                item.InPassQty?.ToString() ?? "",
                                item.InReturnQty?.ToString() ?? "",
                                item.DocDate?.ToString() ?? "",
                                item.SupplierFull?.ToString() ?? "",
                                item.CloseCode?.ToString() ?? "",
                            };

                            rowIndex++;

                            for (int i = 0; i < rowData.Length; i++)
                            {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, rowIndex)).Value = rowData[i];
                            }
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

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
                        worksheet.SheetView.FreezeRows(1);
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


        #region //GetIOTDataExcel IOTEXCEL

        #region //GetIOTDataExcel 氣密
        public void GetIOTDataExcel(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT 氣密報表";
                    string excelsheetName = "IOT 氣密報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "機臺編號", "人員", "製令", "製程", "來料批號條碼", "來料QRCode"
                                                   , "產品作業順序", "來料批號開工時間", "來料批號完工時間", "Machine_CT", "單顆作業時間", "機故時間"
                                                   , "氣密機量測時間", "寫入SQL時間", "氣密壓力源 (Kpa)", "洩漏值(cm3/mn)", "OK/NG", "產品編號"
                                                   , "OK_TrayCode", "NG_TrayCode", "Model"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MachineSN?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.Operater?.ToString().Trim() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MfgOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.Process?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Sup_LotCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Sup_QRCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.StepOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Batch_Start_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Batch_End_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Machine_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Inspection_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Alarm_Duration?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.Inspection_StartTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.RecordTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.Pressure_kpa?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.Leak_Cm3?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.Result?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.Barcode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.OK_TrayCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.NG_TrayCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.Model?.ToString() ?? "";


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel 點膠
        public void GetIOTDataExcel2(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT 點膠報表";
                    string excelsheetName = "IOT 點膠報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "機臺編號", "人員", "製令", "製程", "來料批號條碼", "來料QRCode"
                                                   , "產品作業順序", "來料批號開工時間", "來料批號完工時間", "Machine_CT", "單顆作業時間"
                                                   , "機故時間", "量測時間", "寫入SQL時間", "Spare_Col", "產品編號", "點膠量(g)"
                                                   , "膠材批號", "膠材限制使用時間", "UV固化(S)", "UV照度", "OK/NG", "OK_TrayCode"
                                                   , "NG_TrayCode", "Model"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MachineSN?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.Operater?.ToString().Trim() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MfgOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.Process?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Sup_LotCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Sup_QRCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.StepOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Batch_Start_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Batch_End_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Machine_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Inspection_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Alarm_Duration?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.Inspection_StartTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.RecordTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.SpareCol?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.Barcode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.Glue_g?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.Glue_LotCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.Glue_ExpirationTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.UV_s?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.UV_Illuminance?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.Result?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.OK_TrayCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.NG_TrayCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.Model?.ToString() ?? "";


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel 雷雕
        public void GetIOTDataExcel3(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT 雷雕報表";
                    string excelsheetName = "IOT 雷雕報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "機臺編號", "人員", "製令", "製程", "來料批號條碼", "來料QRCode"
                                                    , "產品作業順序", "來料批號開工時間", "來料批號完工時間", "Machine_CT", "單顆作業時間"
                                                    , "機故時間", "量測時間", "寫入SQL時間", "Spare_Col", "產品編號", "雷雕格式/大小", "等級分類"
                                                    , "OK/NG", "OK_TrayCode", "NG_TrayCode", "Model"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MachineSN?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.Operater?.ToString().Trim() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MfgOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.Process?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Sup_LotCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Sup_QRCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.StepOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Batch_Start_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Batch_End_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Machine_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Inspection_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Alarm_Duration?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.Inspection_StartTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.RecordTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.SpareCol?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.Barcode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.Laser_Size?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.Laser_Class?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.Result?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.OK_TrayCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.NG_TrayCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.Model?.ToString() ?? "";


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel 測高
        public void GetIOTDataExcel4(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT 測高報表";
                    string excelsheetName = "IOT 測高報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[]
                                                    {
                                                        "機臺編號", "人員", "製令", "製程", "來料批號條碼", "來料QRCode",
                                                        "產品作業順序", "來料批號開工時間(同第一顆)", "來料批號完工時間(同最後一顆)", "Machine_CT", "單顆作業時間", "機故時間",
                                                        "測高機量測時間", "寫入SQL時間", "Spare_Col", "產品編號", "MaxHight(mm)", "MinHight(mm)",
                                                        "High(mm)", "OK/NG", "Model", "DelayTime"
                                                    };

                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MachineSN?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.Operater?.ToString().Trim() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MfgOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.Process?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Sup_LotCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Sup_QRCode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.StepOrder?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Batch_Start_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Batch_End_Time?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Machine_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Inspection_CT?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Alarm_Duration?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.Inspection_StartTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.RecordTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.SpareCol?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.Barcode?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.L1_Max_Value?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.L1_Min_Value?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.Measurement_Value?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.Result?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.Model?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.DelayTime?.ToString() ?? "";


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel MTF
        public void GetIOTDataExcel5(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT MTF報表";
                    string excelsheetName = "IOT MTF報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] {
                                                        "量測編號", "機台編號", "人員", "製令", "製程", "來料批號條碼", "來料QRCode", "MTF料盤序號",
                                                        "產品序號", "X座標", "Y座標", "A料盤碼", "B料盤碼", "NG料盤碼", "班別", "抽檢頻率",
                                                        "FBL", "EFL", "W_1S", "W_1T", "W_6S", "W_6T", "W_7S", "W_7T",
                                                        "W_8S", "W_8T", "W_9S", "W_9T", "W_14S", "W_14T", "W_15S", "W_15T",
                                                        "W_16S", "W_16T", "W_17S", "W_17T", "W1S_Peak", "W1T_Peak", "W6S_Peak", "W6T_Peak",
                                                        "W7S_Peak", "W7T_Peak", "W8S_Peak", "W8T_Peak", "W9S_Peak", "W9T_Peak", "W14S_Peak", "W14T_Peak",
                                                        "W15S_Peak", "W15T_Peak", "W16S_Peak", "W16T_Peak", "W17S_Peak", "W17T_Peak", "MTF量測時間", "寫入SQL時間",
                                                        "備用欄位"
                                                    };

                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MeasurementSN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MachineSN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.Operater?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MfgOrder?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Process?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Sup_LotCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.Sup_QRCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.MTF_TraySN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ProductSN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Location_X?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Location_Y?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.A_TrayCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.B_TrayCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.NG_TrayCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.Class?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.CheckFreq?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.FBL?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.EFL?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.W_1S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.W_1T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.W_6S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.W_6T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.W_7S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.W_7T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.W_8S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.W_8T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.W_9S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.W_9T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.W_14S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.W_14T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.W_15S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.W_15T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.W_16S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.W_16T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.W_17S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.W_17T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.W1S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.W1T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.W6S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.W6T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.W7S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.W7T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.W8S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.W8T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.W9S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.W9T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.W14S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(48, rowIndex)).Value = item.W14T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(49, rowIndex)).Value = item.W15S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(50, rowIndex)).Value = item.W15T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(51, rowIndex)).Value = item.W16S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(52, rowIndex)).Value = item.W16T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(53, rowIndex)).Value = item.W17S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(54, rowIndex)).Value = item.W17T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(55, rowIndex)).Value = item.Date_Time?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(56, rowIndex)).Value = item.RecordTime?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(57, rowIndex)).Value = item.Spare_Col?.ToString();

                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel 模造中厚
        public void GetIOTDataExcel6(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT 模造中厚報表";
                    string excelsheetName = "IOT 模造中厚報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[]
                                                    {
                                                        "當日顆數(量測序號)", "穴號", "部番", "厚度值", "厚度檢測結果", "外徑值",
                                                        "外徑檢測結果", "量測人員", "紀錄時間"
                                                    };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.Measurement_SN.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Style.NumberFormat.Format = "@";

                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.HoleSN.ToString();

                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.Model.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.Thickness_Value.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.THK_Result.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.OutsideDiameter_Value.ToString();

                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.OD_Result.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Operator.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.RecordTime.ToString();


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel 紘立MTF
        public void GetIOTDataExcel7(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData3(MachineType, startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】IOT 紘立MTF報表";
                    string excelsheetName = "IOT 紘立MTF報表";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[]
                                                    {
                                                        "ID", "MeasurementSN", "MachineSN", "Operater", "MfgOrder", "Process", "Sup_LotCode", "Sup_QRCode", "MTF_TraySN", "ProductSN",
                                                        "Location_X", "Location_Y", "Class", "CheckFreq", "FBL", "DOF_Index", "DOF_minus", "DOF_plus", "DOF_T", "GDOF_Index",
                                                        "GDOF_minus", "GDOF_plus", "GDOF_T", "EFL", "W_1S", "W_1T", "W_2S", "W_2T", "W_3S", "W_3T",
                                                        "W_4S", "W_4T", "W_5S", "W_5T", "W_6S", "W_6T", "W_7S", "W_7T", "W_8S", "W_8T",
                                                        "W_9S", "W_9T", "W_10S", "W_10T", "W_11S", "W_11T", "W_12S", "W_12T", "W_13S", "W_13T",
                                                        "W_14S", "W_14T", "W_15S", "W_15T", "W_16S", "W_16T", "W_17S", "W_17T", "W1S_Peak", "W1T_Peak",
                                                        "W2S_Peak", "W2T_Peak", "W3S_Peak", "W3T_Peak", "W4S_Peak", "W4T_Peak", "W5S_Peak", "W5T_Peak", "W6S_Peak", "W6T_Peak",
                                                        "W7S_Peak", "W7T_Peak", "W8S_Peak", "W8T_Peak", "W9S_Peak", "W9T_Peak", "W10S_Peak", "W10T_Peak", "W11S_Peak", "W11T_Peak",
                                                        "W12S_Peak", "W12T_Peak", "W13S_Peak", "W13T_Peak", "W14S_Peak", "W14T_Peak", "W15S_Peak", "W15T_Peak", "W16S_Peak", "W16T_Peak",
                                                        "W17S_Peak", "W17T_Peak", "MTF_DateTime", "RecordTime", "Spare_Col"
                                                    };

                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            // 2. Excel 寫入程式碼
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MeasurementSN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MachineSN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.Operater?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MfgOrder?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Process?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Sup_LotCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.Sup_QRCode?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.MTF_TraySN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ProductSN?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.Location_X?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Location_Y?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Class?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.CheckFreq?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.FBL?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.DOF_Index?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.DOF_minus?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.DOF_plus?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.DOF_T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.GDOF_Index?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.GDOF_minus?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.GDOF_plus?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.GDOF_T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.EFL?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.W_1S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.W_1T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.W_2S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.W_2T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.W_3S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.W_3T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.W_4S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.W_4T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.W_5S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.W_5T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.W_6S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.W_6T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.W_7S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.W_7T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.W_8S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.W_8T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.W_9S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.W_9T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.W_10S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.W_10T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.W_11S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.W_11T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.W_12S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.W_12T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(48, rowIndex)).Value = item.W_13S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(49, rowIndex)).Value = item.W_13T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(50, rowIndex)).Value = item.W_14S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(51, rowIndex)).Value = item.W_14T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(52, rowIndex)).Value = item.W_15S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(53, rowIndex)).Value = item.W_15T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(54, rowIndex)).Value = item.W_16S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(55, rowIndex)).Value = item.W_16T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(56, rowIndex)).Value = item.W_17S?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(57, rowIndex)).Value = item.W_17T?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(58, rowIndex)).Value = item.W1S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(59, rowIndex)).Value = item.W1T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(60, rowIndex)).Value = item.W2S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(61, rowIndex)).Value = item.W2T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(62, rowIndex)).Value = item.W3S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(63, rowIndex)).Value = item.W3T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(64, rowIndex)).Value = item.W4S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(65, rowIndex)).Value = item.W4T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(66, rowIndex)).Value = item.W5S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(67, rowIndex)).Value = item.W5T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(68, rowIndex)).Value = item.W6S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(69, rowIndex)).Value = item.W6T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(70, rowIndex)).Value = item.W7S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(71, rowIndex)).Value = item.W7T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(72, rowIndex)).Value = item.W8S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(73, rowIndex)).Value = item.W8T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(74, rowIndex)).Value = item.W9S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(75, rowIndex)).Value = item.W9T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(76, rowIndex)).Value = item.W10S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(77, rowIndex)).Value = item.W10T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(78, rowIndex)).Value = item.W11S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(79, rowIndex)).Value = item.W11T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(80, rowIndex)).Value = item.W12S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(81, rowIndex)).Value = item.W12T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(82, rowIndex)).Value = item.W13S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(83, rowIndex)).Value = item.W13T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(84, rowIndex)).Value = item.W14S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(85, rowIndex)).Value = item.W14T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(86, rowIndex)).Value = item.W15S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(87, rowIndex)).Value = item.W15T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(88, rowIndex)).Value = item.W16S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(89, rowIndex)).Value = item.W16T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(90, rowIndex)).Value = item.W17S_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(91, rowIndex)).Value = item.W17T_Peak?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(92, rowIndex)).Value = item.MTF_DateTime?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(93, rowIndex)).Value = item.RecordTime?.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(94, rowIndex)).Value = item.Spare_Col?.ToString();




                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();




                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel8 紘立C90IOTEXCEL
        public void GetIOTDataExcel8(string MachineType = "", string startDate = "", string endDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTData2("C90_Log", startDate, endDate, "", -1, -1);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】C90缺件查詢";
                    string excelsheetName = "C90缺件查詢";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());

                    string[] header = new string[] { "日期時間", "流程卡號", "QR_Code", "前後蓋", "螺絲", "QR_Code版本"
                                                   , "QR_Code週別", "QR_Code重複", "Dcut_Result", "結果", "RecordTime"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.日期時間?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.流程卡號?.ToString().Trim() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.QR_Code?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.前後蓋?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.螺絲?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.QR_Code版本?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.QR_Code週別?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.QR_Code重複?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Dcut_Result?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.結果?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.RecordTime?.ToString() ?? "";


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();

                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #region //GetIOTDataExcel9 鏡頭條碼IOTEXCEL
        public void GetIOTDataExcel9(string QRNos)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetIOTQRData(QRNos);
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
                    dateStyle.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
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

                    #endregion

                    #region //參數初始化
                    byte[] excelFile;
                    string excelFileName = "【MES2.0】鏡頭條碼IOT資料";
                    string excelsheetName = "鏡頭條碼IOT資料";

                    //dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    //dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["result"].ToString());

                    string[] header = new string[] {
    "QR", "ID", "MeasurementSN", "MachineSN", "Location_X", "Location_Y", "Class", "CheckFreq", "FBL", "DOF_Index",
    "DOF_minus", "DOF_plus", "DOF_T", "GDOF_Index", "GDOF_minus", "GDOF_plus", "GDOF_T", "EFL", "W_1S", "W_1T",
    "W_2S", "W_2T", "W_3S", "W_3T", "W_4S", "W_4T", "W_5S", "W_5T", "W_6S", "W_6T",
    "W_7S", "W_7T", "W_8S", "W_8T", "W_9S", "W_9T", "W_10S", "W_10T", "W_11S", "W_11T",
    "W_12S", "W_12T", "W_13S", "W_13T", "W_14S", "W_14T", "W_15S", "W_15T", "W_16S", "W_16T",
    "W_17S", "W_17T", "W1S_Peak", "W1T_Peak", "W2S_Peak", "W2T_Peak", "W3S_Peak", "W3T_Peak", "W4S_Peak", "W4T_Peak",
    "W5S_Peak", "W5T_Peak", "W6S_Peak", "W6T_Peak", "W7S_Peak", "W7T_Peak", "W8S_Peak", "W8T_Peak", "W9S_Peak", "W9T_Peak",
    "W10S_Peak", "W10T_Peak", "W11S_Peak", "W11T_Peak", "W12S_Peak", "W12T_Peak", "W13S_Peak", "W13T_Peak", "W14S_Peak", "W14T_Peak",
    "W15S_Peak", "W15T_Peak", "W16S_Peak", "W16T_Peak", "W17S_Peak", "W17T_Peak", "MTF_DateTime", "RecordTime", "Status"
};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    int startIndex = 0;
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 20;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 11).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            startIndex = rowIndex + 1;
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.QR?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ID?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MeasurementSN?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MachineSN?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.Location_X?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.Location_Y?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.Class?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.CheckFreq?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.FBL?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.DOF_Index?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.DOF_minus?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.DOF_plus?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.DOF_T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.GDOF_Index?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.GDOF_minus?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.GDOF_plus?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.GDOF_T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.EFL?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.W_1S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.W_1T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.W_2S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.W_2T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.W_3S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.W_3T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.W_4S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.W_4T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.W_5S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.W_5T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.W_6S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.W_6T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.W_7S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.W_7T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(33, rowIndex)).Value = item.W_8S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(34, rowIndex)).Value = item.W_8T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(35, rowIndex)).Value = item.W_9S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(36, rowIndex)).Value = item.W_9T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(37, rowIndex)).Value = item.W_10S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(38, rowIndex)).Value = item.W_10T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(39, rowIndex)).Value = item.W_11S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(40, rowIndex)).Value = item.W_11T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(41, rowIndex)).Value = item.W_12S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(42, rowIndex)).Value = item.W_12T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(43, rowIndex)).Value = item.W_13S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(44, rowIndex)).Value = item.W_13T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(45, rowIndex)).Value = item.W_14S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(46, rowIndex)).Value = item.W_14T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(47, rowIndex)).Value = item.W_15S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(48, rowIndex)).Value = item.W_15T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(49, rowIndex)).Value = item.W_16S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(50, rowIndex)).Value = item.W_16T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(51, rowIndex)).Value = item.W_17S?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(52, rowIndex)).Value = item.W_17T?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(53, rowIndex)).Value = item.W1S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(54, rowIndex)).Value = item.W1T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(55, rowIndex)).Value = item.W2S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(56, rowIndex)).Value = item.W2T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(57, rowIndex)).Value = item.W3S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(58, rowIndex)).Value = item.W3T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(59, rowIndex)).Value = item.W4S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(60, rowIndex)).Value = item.W4T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(61, rowIndex)).Value = item.W5S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(62, rowIndex)).Value = item.W5T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(63, rowIndex)).Value = item.W6S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(64, rowIndex)).Value = item.W6T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(65, rowIndex)).Value = item.W7S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(66, rowIndex)).Value = item.W7T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(67, rowIndex)).Value = item.W8S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(68, rowIndex)).Value = item.W8T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(69, rowIndex)).Value = item.W9S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(70, rowIndex)).Value = item.W9T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(71, rowIndex)).Value = item.W10S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(72, rowIndex)).Value = item.W10T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(73, rowIndex)).Value = item.W11S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(74, rowIndex)).Value = item.W11T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(75, rowIndex)).Value = item.W12S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(76, rowIndex)).Value = item.W12T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(77, rowIndex)).Value = item.W13S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(78, rowIndex)).Value = item.W13T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(79, rowIndex)).Value = item.W14S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(80, rowIndex)).Value = item.W14T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(81, rowIndex)).Value = item.W15S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(82, rowIndex)).Value = item.W15T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(83, rowIndex)).Value = item.W16S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(84, rowIndex)).Value = item.W16T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(85, rowIndex)).Value = item.W17S_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(86, rowIndex)).Value = item.W17T_Peak?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(87, rowIndex)).Value = item.MTF_DateTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(88, rowIndex)).Value = item.RecordTime?.ToString() ?? "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(89, rowIndex)).Value = item.Status?.ToString() ?? "";


                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();

                        #endregion

                        #region //凍結
                        //窗格、首欄、頂端列
                        //worksheet.SheetView.Freeze(1, 1);
                        //worksheet.SheetView.FreezeColumns(1);
                        worksheet.SheetView.FreezeRows(2);
                        #endregion

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

        #endregion




        #endregion

        #region //FOR EIP API
        #region //GetProductionInProcessEIP 取得生產在製資料
        [HttpPost]
        [Route("api/CR/GetProductionInProcess")]
        public void GetProductionInProcessEIP(int ModeId = -1, string WoErpFullNo = "", string MtlItemNo = "", string BarcodeNo = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionInProcessEIP(ModeId, WoErpFullNo, MtlItemNo, BarcodeNo, StartDate, EndDate
                    , OrderBy, PageIndex, PageSize, CustomerIds);
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

        #region //GetWipBarcodeCountDataEIP 取得條碼在製資料筆數
        [HttpPost]
        [Route("api/CR/GetWipBarcodeCountData")]
        public void GetWipBarcodeCountDataEIP(string MoProcessIdListString = "")
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWipBarcodeCountDataEIP(MoProcessIdListString);
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

        #region //GetTimeoutEIP 取得頁面刷新時間
        [HttpPost]
        [Route("api/CR/GetTimeout")]
        public void GetTimeoutEIP(string TypeSchema)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetTimeoutEIP(TypeSchema);
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

        #region //GetQcRecordEIP 取得量測資料
        [HttpPost]
        [Route("api/CR/GetQcRecord")]

        public void GetQcRecordEIP(int QcRecordId = -1, string WoErpFullNo = "", string QcNoticeNo = "", string MtlItemNo = "", string MtlItemName = ""
            , int QcTypeId = -1, string CheckQcMeasureData = "", int UserId = -1, string StartDate = "", string EndDate = "", string BarcodeNo = "", string QcType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcRecordEIP(QcRecordId, WoErpFullNo, QcNoticeNo, MtlItemNo, MtlItemName
                    , QcTypeId, CheckQcMeasureData, UserId, StartDate, EndDate, BarcodeNo, QcType
                    , OrderBy, PageIndex, PageSize, CustomerIds);
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

        #region //GetQcMeasureDataEIP 取得量測詳細資料
        [HttpPost]
        [Route("api/CR/GetQcMeasureData")]

        public void GetQcMeasureDataEIP(int QcRecordId = -1, string BarcodeNo = "", string QcResult = "", string QcItemName = "", string MachineNumber = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcMeasureDataEIP(QcRecordId, BarcodeNo, QcResult, QcItemName, MachineNumber);
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

        #region //GetQcRecordFileForPictureEIP 取得量測記錄檔案(限定PNG格式)
        [HttpPost]
        [Route("api/CR/GetQcRecordFileForPicture")]

        public void GetQcRecordFileForPictureEIP(int QcRecordId = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcRecordFileForPictureEIP(QcRecordId);
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

        #region //GetQcRecordFileEIP 取得量測記錄檔案
        [HttpPost]
        [Route("api/CR/GetQcRecordFile")]

        public void GetQcRecordFileEIP(int QcRecordId = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcRecordFileEIP(QcRecordId);
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

        #region //GetQcBarcodeEIP 取得量測判定資料
        [HttpPost]
        [Route("api/CR/GetQcBarcode")]

        public void GetQcBarcodeEIP(int QcRecordId = -1, string BarcodeNo = "", string QcStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQcBarcodeEIP(QcRecordId, BarcodeNo, QcStatus
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

        #region //GetQtyProductionHistoryEIP 取得數量生產在製資料
        [HttpPost]
        [Route("api/CR/GetQtyProductionHistory")]

        public void GetQtyProductionHistoryEIP(int MoId = -1, int ProcessId = -1, string ProdStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetQryProductionHistoryEIP(MoId, ProcessId, ProdStatus
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

        #region //GetProductionHistoryEIP 取得生產在製資料
        [HttpPost]
        [Route("api/CR/GetProductionHistory")]

        public void GetProductionHistoryEIP(int MoProcessId = -1, string ProdStatus = "", string BarcodeNo = "")
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetProductionHistoryEIP(MoProcessId, ProdStatus, BarcodeNo);
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

        #region //GetBarcodeTrackingEIP 取得生產在製資料
        [HttpPost]
        [Route("api/CR/GetBarcodeTracking")]

        public void GetBarcodeTrackingEIP(int BarcodeId = -1, string BarcodeNo = ""
            , int[] CustomerIds = null, bool unValid = false
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeTrackingEIP(BarcodeId, BarcodeNo
                    , CustomerIds, unValid
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

        #region //GetBarcodeProcessEventEIP 取得條碼過站事件資料
        [HttpPost]
        [Route("api/CR/GetBarcodeProcessEvent")]

        public void GetBarcodeProcessEventEIP(int UserId = -1, string BarcodeNo = "", string ProcessEventName = "", string MtlItemNo = "", string StartDate = "", string FinishDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeProcessEventEIP(UserId, BarcodeNo, ProcessEventName, MtlItemNo, StartDate, FinishDate, OrderBy, PageIndex, PageSize);
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

        #region //GetWoDataEIP --基礎過站資訊
        [HttpPost]
        [Route("api/CR/GetWoData")]
        public void GetWoDataEIP(int ProcessId = -1, string WoErpFullNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", int[] CustomerIds = null, string BarcodeNo = "", string Status = "", string isNewFinish = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWoDataEIP(ProcessId, WoErpFullNo, MtlItemNo, StartDate, EndDate, CustomerIds, BarcodeNo, Status, isNewFinish
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

        #region //GetWipBarcodeEIP 取得特定站別在製條碼
        [HttpPost]
        [Route("api/CR/GetWipBarcode")]
        public void GetWipBarcodeEIP(int MoProcessId = -1, string WipStatus = "", string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetWipBarcodeEIP(MoProcessId, WipStatus, BarcodeNo
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

        #region //API

        #region//業務查詢API
        [HttpPost]
        [Route("api/ERP/SaleRankingListVe2")]
        public void ApiSaleRankingListVe2(string Company, string SecretKey
            , string Type = "", string Model = "", string StartDate = "", string EndDate = ""
            , string CompanyList = "", int PgId = -1, string ExcludeStatus = "",string SaleRankingType="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiSaleRankingListVe2");
                #endregion

                #region //Request
                switch (SaleRankingType) {
                    case "GetSaleRankingListVe1":
                        //取得業務成績排行(BY單據業務) 
                        dataRequest = mesReportDA.GetSaleRankingListVe1(Type, Model, StartDate, EndDate, CompanyList, PgId, ExcludeStatus);
                        break;
                    case "GetSaleRankingListVe2":
                        //取得業務成績排行(BY客戶業務)
                        dataRequest = mesReportDA.GetSaleRankingListVe2(Type, Model, StartDate, EndDate, CompanyList, PgId, ExcludeStatus);
                        break;
                }
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

        #region//ERP財務資料查詢(限ERP)
        [HttpPost]
        [Route("api/ERP/ApiErpView")]
        public void ErpView(string Company, string SecretKey
        , string ViewName = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiErpView");
                #endregion

                #region //Request
                dataRequest = mesReportDA.GetErpView(ViewName, Company);
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

        #region //查詢現況生產異常名單及進行推播程序
        [HttpPost]
        [Route("api/MES/RptMoInformationNotifier")]
        public void RptMoInformationNotifier(string Company = "", string SecretKey = "", string ActualStart = "")
        {
            try
            {
                if (Company.Length <= 0)
                {
                    #region //確認公司別
                    int CurrentCompany = 0;
                    if (Session["CompanySwitch"].ToString() == "manual")
                    {
                        CurrentCompany = Convert.ToInt32(Session["CompanyId"]);
                    }
                    else
                    {
                        CurrentCompany = Convert.ToInt32(Session["UserCompany"]);
                    }
                    #endregion

                    if (CurrentCompany <= 0) throw new SystemException("公司別錯誤!!");
                    dataRequest = basicInformationDA.GetCompany(CurrentCompany, "", "", "A", "", -1, -1);
                    JObject dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            Company = item["CompanyNo"].ToString();
                        }
                    }
                    else
                    {
                        throw new SystemException(dataRequestJson["msg"].ToString());
                    }

                    dataRequest = systemSettingDA.GetApiKey(-1, "", "MES相關報表推播API", "", "", -1, -1);
                    dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            SecretKey = item["KeyText"].ToString();
                        }
                    }
                    else
                    {
                        throw new SystemException(dataRequestJson["msg"].ToString());
                    }
                }

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RptMoInformationNotifier");
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.RptMoInformationNotifier(Company, ActualStart);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
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

        #region //UpdateCompanySwitchForApi 公司別切換
        [HttpPost]
        public void UpdateCompanySwitchForApi(string CompanyNo)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetCompany(-1, CompanyNo, "", "A", "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    if (Convert.ToInt32(Session["CompanyId"]) != Convert.ToInt32(result[0]["CompanyId"]))
                    {
                        Session["CompanyId"] = Convert.ToInt32(result[0]["CompanyId"]);
                        Session["CompanyName"] = result[0]["CompanyName"].ToString();
                        Session["CompanySwitch"] = "manual";
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

            //Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //MoProgressDetailNotInventoryMamoEIP 取得生產在製資料，製令完工未入庫，發mamo訊息通知
        [HttpPost]
        [Route("api/MesReport/MoProgressDetailNotInventoryMamoEIP")]

        public void MoProgressDetailNotInventoryMamoEIP(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "MoProgressDetailNotInventoryMamoEIP");
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.MoProgressDetailNotInventoryMamoEIP(Company);
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

        #region //MrDetailAbnormalEIP 	領料異常推播
        [HttpPost]
        [Route("api/MesReport/MrDetailAbnormalEIP")]

        public void MrDetailAbnormalEIP(string Company,string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "MrDetailAbnormalEIP");
                #endregion

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.MrDetailAbnormalEIP(Company);
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

        #region //階層條碼用
        #region //GetBarcodeStatusInfo -- 取得條碼下階層條碼(FOR LENS)
        [HttpPost]
        public void GetBarcodeStatusInfo(int BarcodeId)
        {
            try
            {
                //WebApiLoginCheck("ProductQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeStatusInfo(BarcodeId);
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



        #region //GetBarcodeStatusInfo2 -- 取得條碼上階層條碼(FOR PACKAGE)
        [HttpPost]
        public void GetBarcodeStatusInfo2(int BarcodeId)
        {
            try
            {
                //WebApiLoginCheck("ProductQuery", "read");

                #region //Request
                mesReportDA = new MesReportDA();
                dataRequest = mesReportDA.GetBarcodeStatusInfo2(BarcodeId);
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

    }
}