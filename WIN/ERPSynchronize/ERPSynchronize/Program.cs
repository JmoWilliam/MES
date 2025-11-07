using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERPSynchronize
{
    class Program
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            try
            {
                string transferType = "";
                string company = "";
                string secretKey = "";
                string channelId = "";
                string sendCustomer = "";
                string typeId = "";


                if (args.Length > 0)
                {
                    transferType = args[0].ToString();
                    company = args[1].ToString();
                    secretKey = args[2].ToString();
                }

                switch (transferType)
                {
                    #region //PDM
                    case "UnitOfMeasureSynchronize":
                        Console.WriteLine("ERP單位資料同步中...");
                        UnitOfMeasureSynchronize unitOfMeasureSynchronize = new UnitOfMeasureSynchronize(company, secretKey);
                        unitOfMeasureSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "MtlItemSynchronize":
                        Console.WriteLine("ERP品號資料同步中...");
                        MtlItemSynchronize mtlItemSynchronize = new MtlItemSynchronize(company, secretKey);
                        mtlItemSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "BomSynchronize":
                        Console.WriteLine("ERP BOM資料同步中...");
                        BomSynchronize bomSynchronize = new BomSynchronize(company, secretKey);
                        bomSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "BomSubstitutionSynchronize":
                        Console.WriteLine("ERP BOM取替代料資料同步中...");
                        BomSubstitutionSynchronize bomSubstitutionSynchronize = new BomSubstitutionSynchronize(company, secretKey);
                        bomSubstitutionSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    #endregion

                    #region //SCM
                    case "CustomerSynchronize":
                        Console.WriteLine("ERP客戶資料同步中...");
                        CustomerSynchronize customerSynchronize = new CustomerSynchronize(company, secretKey);
                        customerSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "InventorySynchronize":
                        Console.WriteLine("ERP庫別資料同步中...");
                        InventorySynchronize inventorySynchronize = new InventorySynchronize(company, secretKey);
                        inventorySynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "SupplierSynchronize":
                        Console.WriteLine("ERP供應商資料同步中...");
                        SupplierSynchronize supplierSynchronize = new SupplierSynchronize(company, secretKey);
                        supplierSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "SaleOrderSynchronize":
                        Console.WriteLine("ERP訂單資料同步中...");
                        SaleOrderSynchronize saleOrderSynchronize = new SaleOrderSynchronize(company, secretKey);
                        saleOrderSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "ExchangeRateSynchronize":
                        Console.WriteLine("ERP匯率資料同步中...");
                        ExchangeRateSynchronize exchangeRateSynchronize = new ExchangeRateSynchronize(company, secretKey);
                        exchangeRateSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "InventoryTransactionSynchronize":
                        Console.WriteLine("ERP庫存異動資料同步中...");
                        InventoryTransactionSynchronize inventoryTransactionSynchronize = new InventoryTransactionSynchronize(company, secretKey);
                        inventoryTransactionSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "TempShippingNoteSynchronize":
                        Console.WriteLine("ERP暫出單資料同步中...");
                        TempShippingNoteSynchronize tempShippingNoteSynchronize = new TempShippingNoteSynchronize(company, secretKey);
                        tempShippingNoteSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "TempShippingReturnNoteSynchronize":
                        Console.WriteLine("ERP暫出歸還單資料同步中...");
                        TempShippingReturnNoteSynchronize tempShippingReturnNoteSynchronize = new TempShippingReturnNoteSynchronize(company, secretKey);
                        tempShippingReturnNoteSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "GoodsReceiptSynchronize":
                        Console.WriteLine("ERP進貨單資料同步中...");
                        GoodsReceiptSynchronize goodsReceiptSynchronize = new GoodsReceiptSynchronize(company, secretKey);
                        goodsReceiptSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "PurchaseOrderSynchronize":
                        Console.WriteLine("ERP採購單資料同步中...");
                        PurchaseOrderSynchronize purchaseOrderSynchronize = new PurchaseOrderSynchronize(company, secretKey);
                        purchaseOrderSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "LotNumberSynchronize":
                        Console.WriteLine("ERP批號資料同步中...");
                        LotNumberSynchronize LotNumberSynchronize = new LotNumberSynchronize(company, secretKey);
                        LotNumberSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "ItemInventorySynchronize":
                        Console.WriteLine("ERP品號庫存資料同步中...");
                        ItemInventorySynchronize ItemInventorySynchronize = new ItemInventorySynchronize(company, secretKey);
                        ItemInventorySynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "PMDSendSynchronize":
                        Console.WriteLine("PMD舜宇報表寄送中...");
                        PMDSendSynchronize PMDSendSynchronize = new PMDSendSynchronize(company, secretKey);
                        PMDSendSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    #endregion

                    #region //MES
                    case "WipOrderSynchronize":
                        Console.WriteLine("ERP 製令資料同步中...");
                        WipOrderSynchronize wipOrderSynchronize = new WipOrderSynchronize(company, secretKey);
                        wipOrderSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "RptMoInformationSynchronize":
                        Console.WriteLine("生產進度異常同步...");
                        RptMoInformationSynchronize rptMoInformationSynchronize = new RptMoInformationSynchronize(company, secretKey);
                        rptMoInformationSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "NLogSynchronize":
                        Console.WriteLine("刪除NLog程序...");
                        NLogSynchronize nLogSynchronize = new NLogSynchronize(company, secretKey);
                        nLogSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "MrDetailAbnormalMAMO":
                        Console.WriteLine("領料異常查詢推播中...");
                        MrDetailAbnormalMAMO mrDetailAbnormalMAMO = new MrDetailAbnormalMAMO(company, secretKey, channelId);
                        mrDetailAbnormalMAMO.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "MoProgressDetailNotInventoryMAMO":
                        Console.WriteLine("製令完工未入庫查詢推播中...");
                        MoProgressDetailNotInventoryMAMO moProgressDetailNotInventoryMAMO = new MoProgressDetailNotInventoryMAMO(company, secretKey, channelId);
                        moProgressDetailNotInventoryMAMO.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "OspAlertMamo":
                        Console.WriteLine("託外逾時未加工MAMO通知推播中...");
                        OspAlertMamo ospAlertMamo = new OspAlertMamo(company, secretKey);
                        ospAlertMamo.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "OspInAlertMamo":
                        Console.WriteLine("託外回廠前MAMO通知推播中...");
                        OspInAlertMamo ospInAlertMamo = new OspInAlertMamo(company, secretKey);
                        ospInAlertMamo.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "UpdateTempMtfQcMeasureData":
                        Console.WriteLine("解析暫存的MTF數據播中...");
                        UpdateTempMtfQcMeasureData updateTempMtfQcMeasureData = new UpdateTempMtfQcMeasureData(company, secretKey);
                        updateTempMtfQcMeasureData.Init();
                        Console.WriteLine("執行完成。");
                        break;

                    case "SendExcelProjectOutCycle":
                        Console.WriteLine("托外進貨週期性報表Mail發送中...");
                        SendExcelProjectOutCycle sendExcelProjectOutCycle = new SendExcelProjectOutCycle(company, secretKey);
                        sendExcelProjectOutCycle.Init();
                        break;

                    case "SendExcelPendingPurchaseRequestReport":
                        Console.WriteLine("已請未購明細(一般＆資產請購)Mail發送中...");
                        SendExcelPendingPurchaseRequestReport sendExcelPendingPurchaseRequestReport = new SendExcelPendingPurchaseRequestReport(company, secretKey);
                        sendExcelPendingPurchaseRequestReport.Init();
                        break;

                    case "SendMWEDetailReportExcel":
                        Console.WriteLine("庫存明細表Mail發送中...");
                        SendMWEDetailReportExcel sendMWEDetailReportExcel = new SendMWEDetailReportExcel(company, secretKey);
                        sendMWEDetailReportExcel.Init();
                        break;

                    case "SendMoMaterialsReportExcel":
                        Console.WriteLine("製令材料在製明細Mail發送中...");
                        SendMoMaterialsReportExcel sendMoMaterialsReportExcel = new SendMoMaterialsReportExcel(company, secretKey);
                        sendMoMaterialsReportExcel.Init();
                        break;
                    case "SendETGMWEDetailReportExcel":
                        Console.WriteLine("庫存明細表Mail發送中...");
                        SendETGMWEDetailReportExcel sendETGMWEDetailReportExcel = new SendETGMWEDetailReportExcel(company, secretKey);
                        sendETGMWEDetailReportExcel.Init();
                        break;
                    case "TxKeyenceProcessFromTemp":
                        Console.WriteLine("進行Keyence過站中...");
                        TxKeyenceProcessFromTemp txKeyenceProcessFromTemp = new TxKeyenceProcessFromTemp(company, secretKey);
                        txKeyenceProcessFromTemp.Init();
                        break;
                    #endregion

                    default:
                        Console.WriteLine("未執行任何程式。");
                        break;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                Console.WriteLine(e.Message);
            }
            finally
            {
                ReadLine(1000);
            }
        }

        static void ReadLine(int millisecond)
        {
            ReadLineDelegate readLineDelegate = Console.ReadLine;
            IAsyncResult asyncResult = readLineDelegate.BeginInvoke(null, null);
            asyncResult.AsyncWaitHandle.WaitOne(millisecond);
            if (asyncResult.IsCompleted)
            {
                string result = readLineDelegate.EndInvoke(asyncResult);
                Console.WriteLine("Read: " + result);
            }
            else
            {
                Console.WriteLine("程式自動關閉");
            }
        }

        delegate string ReadLineDelegate();
    }
}
