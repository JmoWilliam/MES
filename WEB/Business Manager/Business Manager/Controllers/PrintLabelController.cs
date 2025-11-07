using NiceLabel.SDK;
using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;
using System.Reflection;
using NLog;
using MESDA;
using System.Threading;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using SCMDA;

namespace Business_Manager.Controllers
{
    public class PrintLabelController : WebController
    {
        private PrintLabelDA printLabelDA = new PrintLabelDA();
        private ILabel label;
        // GET: PrintLabel
        public ActionResult OutsideIdentificationTag()
        {
            //中潤箱號標籤
            ViewLoginCheck();
            return View();
        }
        public ActionResult InStockLabel()
        {
            //入庫標籤
            ViewLoginCheck();
            return View();
        }

        public ActionResult SunnyCustomerMoldLabel()
        {
            //舜宇出貨標籤
            ViewLoginCheck();
            return View();
        }

        public ActionResult BlackLabel()
        {
            //黑物成型標籤
            ViewLoginCheck();
            return View();
        }

        public ActionResult SunnyBoxLabel()
        {
            //舜宇出貨外箱標籤
            ViewLoginCheck();
            return View();
        }

        public ActionResult LotBarcoderPrintLabel()
        {
            //製令批量條碼標籤
            ViewLoginCheck();
            return View();
        }
        public ActionResult MtlItemPrintLabel()
        {
            //小胖蜂品號標籤列印
            ViewLoginCheck();
            return View();
        }

        public ActionResult MtlItemNoLabel()
        {
            //刀具品號標籤
            ViewLoginCheck();
            return View();
        }
        public ActionResult JCSecondOperationPrintLabel()
        {
            //晶彩二次加工及鍍膜課標籤列印
            ViewLoginCheck();
            return View();
        }

        #region //GetLabelPrintMachine02 取得標籤機資訊
        [HttpPost]
        public void GetLabelPrintMachine02(int LabelPrintId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                dataRequest = printLabelDA.GetLabelPrintMachine02(LabelPrintId
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

        #region //CartonBarcodeprint -- 箱號列印  -- Shintokuro 2022-09-20
        [HttpPost]
        public void CartonBarcodeprint(string CartonBarcode, string PrintMachine)
        {
            try
            {
                string LabelPath = "", line2;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion                

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintCartonBarcode\\Template\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintCartonBarcode\\Template\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印
                label.Variables["CartonBarcode"].SetValue(CartonBarcode);
                label.Print(1);
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
        }
        #endregion        

        #region //PrintToolList -- 工具列印(直接)--工具型號模組-- Shintokuro 2023-01-06
        [HttpPost]
        public void PrintToolList(string ToolList)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";


                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\ToolModelPrint\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\ToolModelPrint\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "cab MACH/300-01";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印

                string ToolNo = "";
                var ToolArr = ToolList.Split(',');

                for (var i = 0; i <= ToolArr.Length; i++)
                {
                    ToolNo = ToolArr[i];
                    label.Variables["ToolNo"].SetValue(ToolArr[i]);
                    label.Print(1);

                }
                #endregion 

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
        }

        #endregion

        #region //PrintToolTemporary -- 工具暫存列表列印--工具型號-- Shintokuro 2023-01-06
        [HttpPost]
        public void PrintToolTemporary(string ToolTemporaryList)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";


                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintCartonBarcode\\Template\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintCartonBarcode\\Template\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "cab MACH/300-01";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印

                string ToolNo = "";
                var ToolArr = ToolTemporaryList.Split(',');

                for (var i = 0; i <= ToolArr.Length; i++)
                {
                    ToolNo = ToolArr[i];
                    label.Variables["ToolNo"].SetValue(ToolArr[i]);
                    //label.Print(1);
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
        }
        #endregion
        #endregion

        #region//PrintQcNoticeNo -- 量測需求單 標籤列印
        [HttpPost]
        public void PrintQcNoticeNo(string QcNoticeNo, string MtlItemNo)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\QcNoticeLabel\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\QcNoticeLabel\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印
                label.Variables["QcNoticeNo"].SetValue(QcNoticeNo);
                label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                label.Print(1);
                #endregion
                
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
        }
        
        #endregion

        #region//替代品標籤
        #endregion

        #region//倉庫

        #region//WarehouseStorageForPcs 入庫標籤

        #region// GetBarcode
        [HttpPost]
        public void GetBarcode(string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("InStockLabel", "read");

                #region //Request                
                dataRequest = printLabelDA.GetBarcode(BarcodeNo);
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

        #region//GetWorkStage --取得加工階段
        [HttpPost]
        public void GetWorkStage(int ModeId = -1)
        {
            try
            {
                WebApiLoginCheck("InStockLabel", "read");

                #region //Request                
                dataRequest = printLabelDA.GetWorkStage(ModeId);
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

        #region//PrintInStockLabel -- 入庫標籤 標籤列印
        [HttpPost]
        public void PrintInStockLabel(string BarcodeNo, string MtlItemNo, string MtlItemName, string Warehouse, string WorkStage)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\LabelData\\InStockLabel\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\LabelData\\InStockLabel\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印
                label.Variables["BARCODE_NO"].SetValue(BarcodeNo);
                label.Variables["MTL_NO_VX"].SetValue(MtlItemNo);
                label.Variables["MTL_NAME_VX"].SetValue(MtlItemName);
                label.Variables["WAREHOUSE"].SetValue(Warehouse);
                label.Variables["STATION"].SetValue(WorkStage);
                label.Print(1);
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
        }
        #endregion
        #endregion

        #endregion

        #region//OutsideIdentificationTag 中潤箱號標籤
        #region //GetLabelPrintMachine --取得標籤機
        [HttpPost]
        public void GetLabelPrintMachine(string LabelPrintNo = "", int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsideIdentificationTag", "read");

                #region //Request                
                dataRequest = printLabelDA.GetLabelPrintMachine(LabelPrintNo, CompanyId);
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

        #region//GetErpCustData 取得產品客戶資訊
        [HttpPost]
        public void GetErpCustData(string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("OutsideIdentificationTag", "read");

                #region //Request                
                dataRequest = printLabelDA.GetBarcode(BarcodeNo);
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

        #region//GetCompany 取得公司
        [HttpPost]
        public void GetLabelCompany(int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsideIdentificationTag", "read,constrained-data");

                #region //Request
                dataRequest = printLabelDA.GetCompany(CompanyId);
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

        #region//OtherCustomerMoldLabel 晶彩&其他 出貨標籤
        #endregion

        #region//SunnyCustomerMoldLabel 舜宇出貨標籤
        #region//GetSunnyCustMoldLabel 
        [HttpPost]
        public void GetSunnyCustMoldLabel(string PrintUser = "", string BarcodeNo = "", int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("SunnyCustomerMoldLabel", "read");

                #region //Request                
                dataRequest = printLabelDA.GetSunnyCustMoldLabel(BarcodeNo, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                foreach (var item in jsonResponse["result"])
                {
                    BarcodeNo = item["BarcodeNo"].ToString();
                    string MtlItemNo = item["BarcodeItem"].ToString();
                    PrintSunnyCustomerMoldLabel(PrintUser, BarcodeNo, MtlItemNo);
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

        #region//PrintSunnyCustomerMoldLabel --列印標籤
        [HttpPost]
        public void PrintSunnyCustomerMoldLabel(string PrintUser, string BarcodeNo, string MtlItemNo)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                if (PrintUser == "1")
                {
                    #region //指定要列印的標籤機與標籤檔案路徑

                    #region //標籤格式圖檔路徑
                    StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\SunnyCustomerMoldLabel\\LabelName.txt", Encoding.Default);
                    while ((line2 = LabelPathTxt.ReadLine()) != null)
                    {
                        LabelPath = line2.ToString();
                    }
                    #endregion

                    #region //印製標籤機路徑
                    if (PrintMachine == "")
                    {
                        StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\SunnyCustomerMoldLabel\\PrinterName.txt", Encoding.Default);
                        while ((line2 = PrintMachineTxt.ReadLine()) != null)
                        {
                            PrintMachine = line2.ToString();
                        }
                        label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                        label.PrintSettings.PrinterName = PrintMachine;
                    }
                    #endregion

                    #region //標籤列印
                    label.Variables["BARCODE_NO"].SetValue(BarcodeNo);
                    label.Variables["WIP_NAME"].SetValue(MtlItemNo);
                    label.Print(1);
                    #endregion
                    #endregion
                }
                else
                {
                    #region //指定要列印的標籤機與標籤檔案路徑

                    #region //標籤格式圖檔路徑
                    StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\SunnyCustomerMoldLabelQa\\LabelName.txt", Encoding.Default);
                    while ((line2 = LabelPathTxt.ReadLine()) != null)
                    {
                        LabelPath = line2.ToString();
                    }
                    #endregion

                    #region //印製標籤機路徑
                    if (PrintMachine == "")
                    {
                        StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\SunnyCustomerMoldLabelQa\\PrinterName.txt", Encoding.Default);
                        while ((line2 = PrintMachineTxt.ReadLine()) != null)
                        {
                            PrintMachine = line2.ToString();
                        }
                        label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                        label.PrintSettings.PrinterName = PrintMachine;
                    }
                    #endregion

                    #region //標籤列印
                    label.Variables["BARCODE_NO"].SetValue(BarcodeNo);
                    label.Variables["WIP_NAME"].SetValue(MtlItemNo);
                    label.Print(1);
                    #endregion
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
        }
        #endregion

        #endregion

        #region//JMOLabel_SUNNY 舜宇箱號標籤        
        #region//GetSunnyCustMoldLabel 
        [HttpPost]
        public void GetSunnyBoxLabel(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("SunnyBoxLabel", "read");

                #region //Request                
                dataRequest = printLabelDA.GetSunnyBoxLabel(MoId);
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

        #region//NewMaxCustomerMoldLabel 新鉅出貨標籤
        #endregion

        #region//printImitationLabel 鎢鋼模仁標籤
        #endregion

        #region//先進光 -- 須向 景瑜確認資料
        #endregion

        #region//替代品標籤
        #endregion

        #endregion        

        #region//成型

        #region//黑物成型標籤
        #region//GetBlackLabel 
        [HttpPost]
        public void GetBlackLabel(string BarcodeNo = "", int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("BlackLabel ", "read");

                #region //Request                
                dataRequest = printLabelDA.GetBlackLabel(BarcodeNo, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                foreach (var item in jsonResponse["result"])
                {
                    BarcodeNo = item["BarcodeNo"].ToString();

                    PrintBlackLabel(BarcodeNo);
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

        #region//PrintBlackLabel --列印標籤
        [HttpPost]
        public void PrintBlackLabel(string BarcodeNo)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\BlackLabel\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\BlackLabel\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印
                label.Variables["BARCODE_NO"].SetValue(BarcodeNo);
                label.Variables["BARCODE_CAVITY"].SetValue(BarcodeNo);
                label.Print(1);
                #endregion
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
        }
        #endregion
        #endregion

        #region//黑物FB成型標籤
        #endregion

        #region //量測
        [HttpPost]
        public void PrintQCLable(int QcRecordId = -1, string PrintMachineName = "")
        {
            try
            {
                #region //取得列印用資料
                //WebApiLoginCheck("BlackLabel ", "read");
                string MOID = "";
                string RID = "";
                string MtlItemName = "";
                string Remark = "";

                string MtlItemNo = "";
                string QcUser = "";
                string QcDepartment = "";

                string QcDate = "";



                #region //Request                
                dataRequest = printLabelDA.GetQCBarcodeLabel(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                foreach (var item in jsonResponse["result"])
                {
                    MOID = item["MOID"].ToString();
                    RID = item["RID"].ToString();
                    MtlItemName = item["MtlItemName"].ToString();
                    Remark = item["Remark"].ToString();
                    MtlItemNo = item["MtlItemNo"].ToString();
                    QcUser = item["QcUser"].ToString();
                    QcDepartment = item["QcDepartment"].ToString();
                    QcDate = item["QcDate"].ToString();
                }
                #endregion

                string LabelPath = "", LabelPathline, PrintMachineline;
                string PrintMachine = PrintMachineName;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\QCLabel\\LabelName.txt", Encoding.Default);
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                    LabelPath = "C:\\WIN\\QCLabel\\QCLable.nlbl";

                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\QCLabel\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion
                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;
                //label.PrintSettings.PrinterName = "1234";
                try
                {
                    #region //標籤列印
                    label.Variables["MOID"].SetValue(MOID);
                    label.Variables["RID"].SetValue(RID);
                    label.Variables["MtlItemName"].SetValue(MtlItemName);
                    label.Variables["Remark"].SetValue(Remark);

                    label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                    label.Variables["QcUser"].SetValue(QcUser);
                    label.Variables["QcDepartment"].SetValue(QcDepartment);
                    label.Variables["QcDate"].SetValue(QcDate);
                    label.Print(1);

                    //try
                    //{
                    //    int copies = 1; // 设定打印份数
                    //    var printJob = label.Print(0);
                    //    Console.WriteLine($"打印任务 ID: {printJob}");
                    //}
                    //catch (Exception ex)
                    //{
                    //    Console.WriteLine("打印发生错误: " + ex.Message);
                    //    Console.WriteLine($"详细错误信息: {ex.StackTrace}");

                    //}

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "列印成功!!"
                    });
                    #endregion
                }
                catch (Exception e)
                {
                    var a = e.ToString();
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = e.ToString()
                    });
                    #endregion

                    logger.Error(e.Message);
                }
                #endregion
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

        #region //一般輸出版本
        public void PrintLotQCLable(string QcRecordIds = "", string PrintMachineName = "")
        {
            try
            {
                #region //取得列印用資料
                //WebApiLoginCheck("BlackLabel ", "read");
                string MOID = "";
                string RID = "";
                string MtlItemName = "";
                string Remark = "";

                string MtlItemNo = "";
                string QcUser = "";
                string QcDepartment = "";
                string QcDate = "";
                string[] QcRecordIdList = QcRecordIds.Split(',');//
                foreach (var qcRecordId in QcRecordIdList)
                {
                    int numqcRecordId = int.Parse(qcRecordId);

                    #region //Request    
                    dataRequest = printLabelDA.GetQCBarcodeLabel(numqcRecordId);
                    #endregion

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    foreach (var item in jsonResponse["result"])
                    {
                        MOID = item["MOID"].ToString();
                        RID = item["RID"].ToString();
                        MtlItemName = item["MtlItemName"].ToString();
                        Remark = item["Remark"].ToString();
                        MtlItemNo = item["MtlItemNo"].ToString();
                        QcUser = item["QcUser"].ToString();
                        QcDepartment = item["QcDepartment"].ToString();
                        QcDate = item["QcDate"].ToString();
                    }

                    string LabelPath = "", LabelPathline, PrintMachineline;
                    string PrintMachine = PrintMachineName;

                    #region //NiceLabel套件
                    string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                    if (Directory.Exists(sdkFilesPath))
                    {
                        PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                    }
                    PrintEngineFactory.PrintEngine.Initialize();
                    #endregion

                    #region //指定要列印的標籤機與標籤檔案路徑

                    #region //標籤格式圖檔路徑
                    StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\QCLabel\\LabelName.txt", Encoding.Default);
                    while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                    {
                        LabelPath = LabelPathline.ToString();
                        LabelPath = "C:\\WIN\\QCLabel\\QCLable.nlbl";

                    }
                    #endregion

                    #region //印製標籤機路徑
                    if (PrintMachine == "")
                    {
                        StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\QCLabel\\PrinterName.txt", Encoding.Default);
                        while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                        {
                            PrintMachine = PrintMachineline.ToString();
                        }
                    }
                    #endregion

                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;

                    try
                    {
                        #region //標籤列印

                        label.Variables["MOID"].SetValue(MOID);
                        label.Variables["RID"].SetValue(RID);
                        label.Variables["MtlItemName"].SetValue(MtlItemName);
                        label.Variables["Remark"].SetValue(Remark);

                        label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                        label.Variables["QcUser"].SetValue(QcUser);
                        label.Variables["QcDepartment"].SetValue(QcDepartment);
                        label.Variables["QcDate"].SetValue(QcDate);
                        label.Print(1);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "列印成功!!"
                        });
                        #endregion

                        #endregion
                    }
                    catch (Exception e)
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = e.ToString()
                        });
                        #endregion

                        logger.Error(e.Message);
                    }
                    #endregion

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

        #region //可定義張數版本

        [HttpPost]
        public void PrintLotQCLable1(string QcRecordObj = "", string PrintMachineName = "")
        {
            try
            {
                if (QcRecordObj.Length < 0) throw new SystemException("標籤資料不可為空");
                if (PrintMachineName.Length < 0) throw new SystemException("標籤機台不可為空");

                #region //取得列印用資料
                //WebApiLoginCheck("BlackLabel ", "read");
                string MOID = "";
                string RID = "";
                string MtlItemName = "";
                string Remark = "";

                string MtlItemNo = "";
                string QcUser = "";
                string QcDepartment = "";
                string QcDate = "";
                string QcRecordIds = "";
                string[] QcRecordIdList = QcRecordIds.Split(',');

                #region //標籤機相關設定
                string LabelPath = "", LabelPathline, PrintMachineline;
                string PrintMachine = PrintMachineName;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\QCLabel\\LabelName.txt", Encoding.Default);
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                    LabelPath = "C:\\WIN\\QCLabel\\QCLable.nlbl";

                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\QCLabel\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion

                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;
                #endregion

                List<Dictionary<string, string>> QcRecordObjJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(QcRecordObj);
                string timeStr = "";
                foreach(var item in QcRecordObjJsonList)
                {

                    int recordId = Convert.ToInt32(item["RecordId"]);
                    int printQty = Convert.ToInt32(item["PrintQty"]);
                    string printDate = item["PrintDate"];
                    if(printQty <1) throw new SystemException("列印數量不可以小於1");
                    if(recordId < 0) throw new SystemException("量測單據不可為空");

                    #region //Request    
                    dataRequest = printLabelDA.GetQCBarcodeLabel(recordId);
                    #endregion

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    foreach (var item1 in jsonResponse["result"])
                    {
                        MOID = item1["MOID"].ToString();
                        RID = item1["RID"].ToString();
                        MtlItemName = item1["MtlItemName"].ToString();
                        Remark = item1["Remark"].ToString();
                        MtlItemNo = item1["MtlItemNo"].ToString();
                        QcUser = item1["QcUser"].ToString();
                        QcDepartment = item1["QcDepartment"].ToString();
                        QcDate = item1["QcDate"].ToString();
                    }

                    #region //標籤列印
                    try
                    {
                        label.Variables["MOID"].SetValue(MOID);
                        label.Variables["RID"].SetValue(RID);
                        label.Variables["MtlItemName"].SetValue(MtlItemName);
                        label.Variables["Remark"].SetValue(Remark);

                        label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                        label.Variables["QcUser"].SetValue(QcUser);
                        label.Variables["QcDepartment"].SetValue(QcDepartment);
                        label.Variables["QcDate"].SetValue(printDate !="" ? printDate : QcDate);
                        for(int num = 1; num <= printQty; num++)
                        {
                            timeStr += $"量測單{RID}:印{num}";
                            label.Print(1);                       
                        }
                    }
                    catch (Exception e)
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = e.ToString()
                        });
                        #endregion

                        logger.Error(e.Message);
                    }
                    #endregion

                }


                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = $"列印成功!!{timeStr}"
                });
                #endregion

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

        #endregion

        #region//模造
        #endregion

        #region//組立
        #endregion

        #region //InitializePrintEngine
        private void InitializePrintEngine()
        {
            try
            {
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
            }
            catch (SDKException exception)
            {
                logger.Error("Initialization of the SDK failed." + Environment.NewLine + Environment.NewLine + exception.ToString());
            }
        }
        #endregion

        #region //包裝標籤列印
        public void GetPackageLabel(int PackageBarcodeId = -1, string PrintMachineName = "")
        {
            try
            {
                WebApiLoginCheck("BlackLabel ", "read");
                string PackageBarcodeNo = "";
                string TotalCount = "";

                #region //Request                
                dataRequest = printLabelDA.GetPackageBarcodeLabel(PackageBarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                foreach (var item in jsonResponse["result"])
                {
                    PackageBarcodeNo = item["BarcodeNo"].ToString();
                    TotalCount = item["TotalCount"].ToString();

                    PrintPackageBarcode(PackageBarcodeNo, PrintMachineName, TotalCount);
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

        #region//PrintPackageBarcode --列印包裝標籤
        [HttpPost]
        public void PrintPackageBarcode(string PackageBarcodeNo, string PrintMachineName, string TotalCount)
        {
            try
            {
                string LabelPath = "", line2;
                string PrintMachine = "";

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PackageBarcode\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑PrintMachineName
                PrintMachine = PrintMachineName;
                //PrintMachine = "Qirui QR-668E";
                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PackageBarcode\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "POSTEK C168/ 300s";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印
                label.Variables["PackageBarcode_No"].SetValue(PackageBarcodeNo);
                label.Variables["PackageBarcode_Count"].SetValue(TotalCount);
                label.Print(1);
                #endregion
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
        }
        #endregion

        #region//PrintNewLabel --空條碼列印
        [HttpPost]
        public void PrintNewLabel(string NewBarcodeNo, string PrintMachine)
        {
            try
            {
                string LabelPath = "", LabelPathline, PrintMachineline;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\NewBarcodeNo\\LabelName.txt", Encoding.Default);
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\NewBarcodeNo\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion

                #region//取得班別
                string Shift = "";
                printLabelDA = new PrintLabelDA();
                dataRequest = printLabelDA.GetBarcodeAttribute(NewBarcodeNo, "Shift");
                var status = JObject.Parse(dataRequest)["status"];
                if (JObject.Parse(dataRequest)["status"].ToString() == "null")
                {
                    Shift = "";
                }
                else
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    Shift = result[0]["ItemValue"].ToString();
                }
                #endregion


                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;
                #region //標籤列印
                label.Variables["BarcodeNo"].SetValue(NewBarcodeNo);
                label.Variables["BarcodeNoShift"].SetValue(NewBarcodeNo + " " + Shift);
                label.Print(1);
                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK"
                });
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

        #region//PrintBindBarcodeNo --數量產條碼列印
        [HttpPost]
        public void PrintBindBarcodeNo(string BarcodeNoList = "", string PrintMachine = "")
        {
            try
            {
                string LabelPath = "", LabelPathline, PrintMachineline;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\NewBarcodeNo\\LabelName.txt", Encoding.Default);
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\NewBarcodeNo\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion

                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;

                #region //標籤列印
                string[] BarcodeNoList2 = BarcodeNoList.Split(',');
                foreach (var barcodeNo in BarcodeNoList2)
                {
                    #region//取得班別
                    string Shift = "";
                    printLabelDA = new PrintLabelDA();
                    dataRequest = printLabelDA.GetBarcodeAttribute(barcodeNo, "Shift");
                    var status = JObject.Parse(dataRequest)["status"];
                    if (JObject.Parse(dataRequest)["status"].ToString() == "null")
                    {
                        Shift = "";
                    }
                    else
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        Shift = result[0]["ItemValue"].ToString();

                    }
                    #endregion
                    label.Variables["BarcodeNoShift"].SetValue(barcodeNo + " " + Shift);
                    label.Variables["BarcodeNo"].SetValue(barcodeNo);
                    label.Print(1);
                }
                #endregion

                #region //成功列印後，更改條碼數量
                dataRequest = printLabelDA.UpdateBarcodePrintCount(BarcodeNoList);
                #endregion

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

        #region //PrintLotBarcode 列印批量條碼
        [HttpPost]
        public void PrintLotBarcode(string BarcodeNoList = "", string PrintMachine = "", string MtlItemName = "", string PrintType = "", int MoId = -1)
        {
            try
            {
                string LabelPath = "", LabelPathline, PrintMachineline;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                string CurrentLabelPath = "";
                if (PrintMachine == "CAB MACH/200-01")
                {
                    CurrentLabelPath = "C:\\WIN\\BarcodeAttribute\\LabelNameFor200.txt";
                }
                else
                {
                    CurrentLabelPath = "C:\\WIN\\BarcodeAttribute\\LabelName.txt";
                }
                StreamReader LabelPathTxt = new StreamReader(CurrentLabelPath, Encoding.Default);
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\BarcodeAttribute\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion

                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;

                #region //處理品名中文問題
                string pattern = @"[\u4e00-\u9fa5]";
                MtlItemName = Regex.Replace(MtlItemName, pattern, "");
                #endregion

                if (PrintType == "section")
                {
                    #region //確認條碼資料是否正確
                    dataRequest = printLabelDA.CheckBarcodePrint(BarcodeNoList);
                    JObject dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["status"].ToString() != "success")
                    {
                        throw new SystemException(dataRequestJson["msg"].ToString());
                    }
                    #endregion

                    #region //標籤列印
                    string[] BarcodeNoList2 = BarcodeNoList.Split(',');//
                    foreach (var barcodeNo in BarcodeNoList2)
                    {
                        label.Variables["BarcodeNo"].SetValue(barcodeNo);
                        label.Variables["MtlItemName"].SetValue(MtlItemName);
                        label.Print(1);
                        Thread.Sleep(2000);
                    }
                    #endregion
                }
                else if (PrintType == "all")
                {
                    #region //取得此製令下所有批量條碼
                    dataRequest = printLabelDA.GetBarcodePrint(MoId);
                    JObject dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["status"].ToString() != "success")
                    {
                        throw new SystemException(dataRequestJson["msg"].ToString());
                    }
                    else
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                label.Variables["BarcodeNo"].SetValue(item["BarcodeNo"].ToString());
                                label.Variables["MtlItemName"].SetValue(MtlItemName);
                                label.Print(1);
                                Thread.Sleep(2000);
                            }
                        }
                        else
                        {
                            throw new SystemException("查無此製令條碼資料!!");
                        }
                    }
                    #endregion
                }
                else
                {
                    throw new SystemException("條碼列印範圍錯誤，請重新刷新頁面後再嘗試。");
                }

                #region //成功列印後，更改條碼列印數量
                dataRequest = printLabelDA.UpdateBarcodePrintCount(BarcodeNoList);
                #endregion

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

        #region //MoNoBarcodePrint -- 批量條碼列印  -- Shintokuro 2023-06-01
        [HttpPost]
        public void MoNoBarcodePrint(string MoId = "", string BarcodeNo = "", string PrintMachine = "", string PrintCnt = "", string TemporaryBarcodeList = "", string ViewCompanyId = "", string LabelFormat = "")
        {
            try
            {
                WebApiLoginCheck("LotBarcoderPrintLabel", "print");
                var printers = PrintEngineFactory.PrintEngine.Printers;

                int PrintCout = 0;
                string BarcodeNoStr = "";
                string LabelPath = "", line2;
                //PrintMachine = "TSC TTP-345";
                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                try
                {
                    PrintEngineFactory.PrintEngine.Initialize();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("SDK 初始化失敗: " + ex.Message);
                }

                PrintEngineFactory.PrintEngine.Initialize();
                #endregion


                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                if (PrintMachine == "ZDesigner ZT410-300dpi ZPL-02")
                {
                    LabelPath = "C:\\WIN\\PrintLotBarcode\\Template\\LotBarcode_BLACK.nlbl";
                }
                else
                {
                    if (ViewCompanyId == "4")
                    {
                        StreamReader LabelPathTxt = null;
                        if (LabelFormat == "A")
                        {
                            LabelPathTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\LotBarcodeDG.txt", Encoding.Default); ;

                        }
                        else if (LabelFormat == "B")
                        {
                            LabelPathTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\ProcessCardDG.txt", Encoding.Default);
                        }
                        else if (LabelFormat == "C")
                        {
                            LabelPathTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\PackageDG.txt", Encoding.Default);
                        }
                        else
                        {
                            LabelPathTxt = new StreamReader("C:\\WIN\\PrintLotBarcode\\Template\\LabelNameDG.txt", Encoding.Default);
                        }

                        while ((line2 = LabelPathTxt.ReadLine()) != null)
                        {
                            LabelPath = line2.ToString();
                        }
                    }
                    else
                    {
                        StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintLotBarcode\\Template\\LabelName.txt", Encoding.Default);
                        while ((line2 = LabelPathTxt.ReadLine()) != null)
                        {
                            LabelPath = line2.ToString();
                        }
                    }
                }
                #endregion

                #region //印製標籤機路徑 //單機版標籤機無下拉會用到
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintLotBarcode\\Template\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "cab MACH/300-01";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                else
                {
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion
                #endregion

                #region //資料取得                
                if (MoId.Length > 0) //製令列印
                {
                    dataRequest = printLabelDA.GetLotBarcodeForLabel(MoId, "", "0", "");
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        for (var i = 0; i < result.Count(); i++)
                        {
                            #region //標籤程式賦直列印
                            MoId = result[i]["MoId"].ToString(); //原MoId有可能是大字串　透過這邊轉換數字Id
                            string barcdoeNo = result[i]["BarcodeNo"].ToString();
                            string WoErpFull = result[i]["WoErpFull"].ToString();
                            string MtlItemName = result[i]["MtlItemName"].ToString();
                            label.Variables["BarcodeNo"].SetValue(barcdoeNo);
                            label.Variables["WoErpFull"].SetValue(WoErpFull);
                            label.Variables["MtlItemName"].SetValue(MtlItemName);

                            string ItemValue = "";
                            JObject jsonObject = JObject.Parse(result[i].ToString());
                            bool hasItemValue = jsonObject.ContainsKey("ItemValue");
                            if (hasItemValue)
                            {
                                ItemValue = result[i]["ItemValue"].ToString();
                                label.Variables["Cavity"].SetValue("Cavity:" + ItemValue);
                            }

                            // 等待两秒钟
                            Thread.Sleep(2000);
                            //try
                            //{
                            //    int copies = 1; // 设定打印份数
                            //    var printJob = label.Print(copies);
                            //    Console.WriteLine($"打印任务 ID: {printJob}");
                            //}
                            //catch (Exception ex)
                            //{
                            //    Console.WriteLine("打印发生错误: " + ex.Message);
                            //    Console.WriteLine($"详细错误信息: {ex.StackTrace}");

                            //}
                            label.Print(1);
                            #endregion
                        }

                        #region //更新條碼列印次數
                        #region //傳送資料到資料庫
                        dataRequest = printLabelDA.UpdateLotPrintCnt("PM", Convert.ToInt32(MoId), "", "");
                        #endregion
                        #region //資料回傳
                        jsonResponse = BaseHelper.DAResponse(dataRequest);
                        #endregion
                        if (jsonResponse.First.First.ToString() == "errorForDA")
                        {
                            throw new SystemException(jsonResponse.Last.Last.ToString());
                        }
                        #endregion
                    }
                    #endregion
                }
                if (BarcodeNo.Length > 0) //條碼列印
                {
                    #region //傳送資料到資料庫
                    dataRequest = printLabelDA.GetLotBarcodeForLabel("", BarcodeNo, "0", "");
                    #endregion

                    #region //資料回傳
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        for (var i = 0; i < result.Count(); i++)
                        {
                            #region //標籤程式賦直列印

                            string barcdoeNo = result[i]["BarcodeNo"].ToString();
                            string WoErpFull = result[i]["WoErpFull"].ToString();
                            string MtlItemName = result[i]["MtlItemName"].ToString();
                            label.Variables["BarcodeNo"].SetValue(barcdoeNo);
                            label.Variables["WoErpFull"].SetValue(WoErpFull);
                            label.Variables["MtlItemName"].SetValue(MtlItemName);

                            string ItemValue = "";
                            JObject jsonObject = JObject.Parse(result[i].ToString());
                            bool hasItemValue = jsonObject.ContainsKey("ItemValue");

                            if (hasItemValue)
                            {
                                ItemValue = result[i]["ItemValue"].ToString();
                                label.Variables["Cavity"].SetValue("Cavity:" + ItemValue);
                            }
                            //try
                            //{
                            //    int copies = 1; // 设定打印份数
                            //    var printJob = label.Print(copies);
                            //    Console.WriteLine($"打印任务 ID: {printJob}");
                            //}
                            //catch (Exception ex)
                            //{
                            //    Console.WriteLine("打印发生错误: " + ex.Message);
                            //    Console.WriteLine($"详细错误信息: {ex.StackTrace}");

                            //}
                            label.Print(1);
                            #endregion

                            #region //更新條碼列印次數
                            dataRequest = printLabelDA.UpdateLotPrintCnt("PB", -1, barcdoeNo, "");
                            #region //Response
                            jsonResponse = BaseHelper.DAResponse(dataRequest);
                            #endregion
                            if (jsonResponse.First.First.ToString() == "errorForDA")
                            {
                                break;
                            }
                            #endregion
                        }
                    }
                    #endregion
                }
                if (TemporaryBarcodeList.Length > 0) //暫存列表列印
                {
                    string TemporaryBarcodeNo = "";
                    var TemporaryBarcodeArr = TemporaryBarcodeList.Split(',');

                    for (var i = 0; i < TemporaryBarcodeArr.Length; i++)
                    {
                        TemporaryBarcodeNo = TemporaryBarcodeArr[i];
                        #region //傳送資料到資料庫
                        dataRequest = printLabelDA.GetLotBarcodeForLabel("", TemporaryBarcodeNo, "0", "");
                        #endregion

                        #region //資料回傳
                        jsonResponse = BaseHelper.DAResponse(dataRequest);
                        if (jsonResponse["status"].ToString() == "success")
                        {
                            string barcdoeNo = "";
                            var result = JObject.Parse(dataRequest)["data"];
                            if (result.Count() <= 0) throw new SystemException("找不到資料,請確認該製令生產模式是否為黑物加工,其批量條碼是否有維護穴號!!");

                            string moId = result[0]["MoId"].ToString();
                            barcdoeNo = result[0]["BarcodeNo"].ToString();
                            string WoErpFull = result[0]["WoErpFull"].ToString();
                            string MtlItemName = result[0]["MtlItemName"].ToString();
                            string planQty = result[0]["PlanQty"].ToString();
                            switch (LabelFormat)
                            {
                                case "A":
                                    label.Variables["MoId"].SetValue(moId);
                                    label.Variables["BarcodeNo"].SetValue(barcdoeNo);
                                    label.Variables["WoErpFull"].SetValue(WoErpFull);
                                    label.Variables["MtlItemName"].SetValue(MtlItemName);

                                    // 等待两秒钟
                                    Thread.Sleep(2000);
                                    label.Print(1);
                                    break;

                                case "B":
                                    label.Variables["MoId"].SetValue(moId);
                                    label.Variables["BarcodeNo"].SetValue(barcdoeNo);
                                    label.Variables["WoErpFull"].SetValue(WoErpFull);
                                    label.Variables["MtlItemName"].SetValue(MtlItemName);
                                    label.Variables["PlanQty"].SetValue(planQty);
                                    Thread.Sleep(2000);
                                    label.Print(1);
                                    break;

                                default:
                                    #region //標籤程式賦直列印
                                    
                                    label.Variables["BarcodeNo"].SetValue(barcdoeNo);
                                    label.Variables["WoErpFull"].SetValue(WoErpFull);
                                    label.Variables["MtlItemName"].SetValue(MtlItemName);

                                    string ItemValue = "";
                                    JObject jsonObject = JObject.Parse(result[0].ToString());
                                    bool hasItemValue = jsonObject.ContainsKey("ItemValue");

                                    if (hasItemValue)
                                    {
                                        ItemValue = result[0]["ItemValue"].ToString();
                                        label.Variables["Cavity"].SetValue("Cavity:" + ItemValue);
                                    }

                                    // 等待两秒钟
                                    Thread.Sleep(2000);
                                    label.Print(1);
                                    #endregion

                                    break;
                            }

                            #region //儲存已經列印的條碼
                            if (PrintCout == 0)
                            {
                                BarcodeNoStr += barcdoeNo;
                            }
                            else
                            {
                                BarcodeNoStr += "," + barcdoeNo;
                            }
                            PrintCout++;
                            #endregion
                        }
                        #endregion
                    }
                    #region //更新條碼列印次數
                    #region //傳送資料到資料庫
                    dataRequest = printLabelDA.UpdateLotPrintCnt("PT", -1, "", BarcodeNoStr);
                    #endregion
                    #region //資料回傳
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
                    if (jsonResponse.First.First.ToString() == "errorForDA")
                    {
                        throw new SystemException(jsonResponse.Last.Last.ToString());
                    }
                    #endregion
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

        #region //UpdateLotPrintCnt 更新批量條碼列印次數
        [HttpPost]
        public void UpdateLotPrintCnt(string PrintType = "", int MoId = -1, string BarcodeNo = "", string BarcodeNoStr = "")
        {
            try
            {
                WebApiLoginCheck("LotBarcoderPrintLabel", "print");

                #region //Request
                dataRequest = printLabelDA.UpdateLotPrintCnt(PrintType, MoId, BarcodeNo, BarcodeNoStr);
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

        #region //UpdateBarcodePrintCount 更新條碼列印次數
        [HttpPost]
        public void UpdateBarcodePrintCount(string BarcodeNoList = "")
        {
            try
            {
                #region //Request
                printLabelDA = new PrintLabelDA();
                dataRequest = printLabelDA.UpdateBarcodePrintCount(BarcodeNoList);
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

        #region //AddBarcodePrint 新增批量條碼
        [HttpPost]
        public void AddBarcodePrint(string WoErpFullNo = "", int BarcodeQty = -1, int BarcodeCount = -1)
        {
            try
            {
                WebApiLoginCheck("LotBarcoderPrintLabel", "add");

                #region //Request
                printLabelDA = new PrintLabelDA();
                dataRequest = printLabelDA.AddBarcodePrint(WoErpFullNo, BarcodeQty, BarcodeCount);
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

        #region//小胖蜂品號標籤列印

        #endregion

        #region //刀具品號標籤列印
        #region //GetMtlItem --取得品號
        [HttpPost]
        public void GetMtlItem(string MtlItemNo = "", string MtlItemNoList = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                printLabelDA = new PrintLabelDA();
                dataRequest = printLabelDA.GetMtlItem(MtlItemNo, MtlItemNoList);
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

        #region //PrintMtlItem -- 品號列印--刀具(兩排的標籤)-- Shintokuro 2023-03-11
        [HttpPost]
        public void PrintMtlItem(string PrintMachine, string MtlItemList)
        {
            try
            {
                string LabelPath = "", line2;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintKnife\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintKnife\\PrintKnife.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "cab MACH/300-01";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                else
                {
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #endregion

                #region //資料解析提取
                var MtlItemArr = MtlItemList.Split(',');
                List<string> PrintMtlItemList = new List<string>();
                for (var i = 0; i < MtlItemArr.Length; i++)
                {
                    int PrintQty = Convert.ToInt32(MtlItemArr[i].Split('❤')[0]);
                    string MtlItemNo = MtlItemArr[i].Split('❤')[1];
                    string MtlItemName = MtlItemArr[i].Split('❤')[2];
                    string MtlItemSpec = MtlItemArr[i].Split('❤')[3];

                    for (var n = 1; n <= PrintQty; n++)
                    {
                        PrintMtlItemList.Add(MtlItemNo + "❤" + MtlItemName + "❤" + MtlItemSpec);
                    }
                }
                #endregion

                #region //標籤列印
                int itemQty = 1;
                int LastTime = PrintMtlItemList.Count();
                int nowTime = 0;
                string MtlItemNo1 = "";
                string MtlItemNo2 = "";
                foreach (var item in PrintMtlItemList)
                {
                    nowTime++;
                    if (itemQty == 1)
                    {
                        MtlItemNo1 = item;
                        label.Variables["MtlItemNo1"].SetValue(item.Split('❤')[0]);
                        label.Variables["MtlItemName1"].SetValue(item.Split('❤')[1]);
                        label.Variables["MtlItemSpec1"].SetValue(item.Split('❤')[2]);
                        itemQty++;
                        if (nowTime == LastTime)
                        {
                            label.Variables["MtlItemNo2"].SetValue("");
                            label.Variables["MtlItemName2"].SetValue("");
                            label.Variables["MtlItemSpec2"].SetValue("");
                            // 等待两秒钟
                            Thread.Sleep(2000);
                            label.Print(1);
                        }
                    }
                    else if (itemQty == 2)
                    {
                        MtlItemNo2 = item;
                        label.Variables["MtlItemNo2"].SetValue(item.Split('❤')[0]);
                        label.Variables["MtlItemName2"].SetValue(item.Split('❤')[1]);
                        label.Variables["MtlItemSpec2"].SetValue(item.Split('❤')[2]);
                        // 等待两秒钟
                        Thread.Sleep(2000);
                        label.Print(1);
                        itemQty = 1;
                    }
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "列印成功"
                });
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

        #region //批號條碼標籤列印
        [HttpPost]
        public void PrintLotNumBarcode(int LotNumberId = -1, string LnBarcodeNoList = "", string PrintMachine = "")
        {
            try
            {
                string LabelPath = "", LabelPathline, PrintMachineline;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\BarcodeAttribute\\LabelName.txt", Encoding.Default);
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\BarcodeAttribute\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion

                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;

                #region //確認條碼資料是否正確
                dataRequest = printLabelDA.GetLotLabel(LotNumberId, LnBarcodeNoList);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //標籤列印
                string[] BarcodeNoList2 = LnBarcodeNoList.Split(',');
                foreach (var barcodeNo in BarcodeNoList2)
                {
                    label.Variables["BarcodeNo"].SetValue(barcodeNo);
                    label.Print(1);
                }
                #endregion

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

        #region //GetJCSecondOperationLabel 晶彩二次加工及鍍膜課多種標籤列印
        [HttpPost]
        public void GetJCSecondOperationLabel(string PrintMachine, string BarcodeNo = "", int MoId = -1, string LabelType = "")
        {
            //LabelType A = 上蓋標籤
            //LabelType B = 流程標籤
            //LabelType C = 包裝標籤

            try
            {
                WebApiLoginCheck("JCSecondOperationPrintLabel ", "read");
                string LabelPath = "", LabelPathline, PrintMachineline;
                List<string> barcodeList = new List<string>();
                List<string> mtlItemNoList = new List<string>();
                List<string> mtlItemNameList = new List<string>();
                List<string> moIdList = new List<string>();
                //if (LabelType == "B")
                {
                    if (MoId <= 0)
                        throw new Exception("請輸入製令條碼！");
                }

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = LabelPathTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\LotBarcodeDG.txt", Encoding.Default); ;
                if (LabelType == "B")
                {
                    LabelPathTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\ProcessCardDG.txt", Encoding.Default);
                }
                else if (LabelType == "C")
                {
                    LabelPathTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\PackageDG.txt", Encoding.Default);
                }
                while ((LabelPathline = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = LabelPathline.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\JCSecondOperation\\PrinterName.txt", Encoding.Default);
                    while ((PrintMachineline = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = PrintMachineline.ToString();
                    }
                }
                #endregion

                #endregion

                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                label.PrintSettings.PrinterName = PrintMachine;
                string BarcodeNoList = "";

                #region //取得條碼資料
                if (BarcodeNo == "" && MoId <= 0)
                {
                    throw new SystemException("產品條碼與製令不可以同時為空值");
                }

                dataRequest = printLabelDA.GetBarcodeandOrder(BarcodeNo, MoId);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    if (dataRequestJson["data"].Count() <= 0)
                    {
                        throw new SystemException("找不到條碼，請確認是否有產生批量條碼！");
                    }
                    #region //列印條碼
                    switch (LabelType)
                    {
                        //Cavity, PotNumber, Shift
                        case "A":
                            #region //上蓋標籤列印
                            foreach (var item in dataRequestJson["data"])
                            {
                                string barcode = item["BarcodeNo"].ToString();
                                if (barcode != "")
                                {
                                    if (BarcodeNoList != "") BarcodeNoList += ",";
                                    string mtlitemName = item["MtlItemName"].ToString();
                                    string moId = item["MoId"].ToString();
                                    string woErpFull = item["WoErpFull"].ToString();

                                    label.Variables["BarcodeNo"].SetValue(barcode);
                                    label.Variables["MtlItemName"].SetValue(mtlitemName);
                                    label.Variables["MoId"].SetValue(moId);
                                    label.Variables["WoErpFull"].SetValue(woErpFull);
                                    label.Print(1);
                                    BarcodeNoList += barcode;
                                }
                            }
                            #endregion

                            #region //成功列印後，更改條碼數量
                            dataRequest = printLabelDA.UpdateBarcodePrintCount(BarcodeNoList);
                            #endregion
                            break;
                        case "B":
                            #region //流程卡標籤列印
                            foreach (var item in dataRequestJson["data"])
                            {
                                string barcode = item["BarcodeNo"].ToString();
                                string mtlitemName = item["MtlItemName"].ToString();
                                string moId = item["MoId"].ToString();
                                string planQty = item["PlanQty"].ToString();
                                string woErpFull = item["WoErpFull"].ToString();

                                label.Variables["BarcodeNo"].SetValue(barcode);
                                label.Variables["PlanQty"].SetValue(planQty);
                                label.Variables["MtlItemName"].SetValue(mtlitemName);
                                label.Variables["MoId"].SetValue(moId);
                                label.Variables["WoErpFull"].SetValue(woErpFull);
                                label.Print(1);
                                break;
                            }
                            #endregion
                            break;
                        default:
                            throw new SystemException("條碼列印範圍錯誤，請重新刷新頁面後再嘗試。");
                    }
                }
                #endregion
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
    }
}