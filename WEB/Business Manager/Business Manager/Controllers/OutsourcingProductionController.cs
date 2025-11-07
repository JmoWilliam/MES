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
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using ZXing;
using ZXing.Common;

namespace Business_Manager.Controllers
{
    public class OutsourcingProductionController : WebController
    {
        private OutsourcingProductionDA outsourcingProductionDA = new OutsourcingProductionDA();
        private ProductionHistoryDA productionHistoryDA = new ProductionHistoryDA();

        #region //View
        public ActionResult OutsourcingProduction()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult OspReceipt()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult OspReceiptConfirm()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetOutsourcingProduction 取得託外生產單資料
        [HttpPost]
        public void GetOutsourcingProduction(int OspId = -1, int DepartmentId = -1, string OspNo = "", int SupplierId = -1, string OspStatus ="", string Status = "", string OspStartDate = "", string OspEndDate = "", string WoErpFullNo = "", string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "read,constrained-data");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOutsourcingProduction(OspId, DepartmentId, OspNo, SupplierId, OspStatus, Status, OspStartDate, OspEndDate, WoErpFullNo, MtlItemNo
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

        #region //GetOspDetail 取得託外生產單詳細資料
        [HttpPost]
        public void GetOspDetail(int OspDetailId = -1, int OspId = -1, int MoId = -1, string Status = "", int SupplierId = -1, string WoErpFullNo = "", string OspNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "read,constrained-data");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspDetail(OspDetailId, OspId, MoId, Status, SupplierId, WoErpFullNo, OspNo
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

        #region //GetOspDetailForPop 取得託外生產單詳細資料(For Pop)
        [HttpPost]
        public void GetOspDetailForPop(int OspDetailId = -1, int OspId = -1, int MoId = -1, string Status = ""
            , int SupplierId = -1, string WoErpFullNo = "", string OspNo = "", string ProcessAlias = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "read,constrained-data");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspDetailForPop(OspDetailId, OspId, MoId, Status
                    , SupplierId, WoErpFullNo, OspNo, ProcessAlias
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

        #region //GetBarcode 取得製令條碼資料
        [HttpPost]
        public void GetBarcode(int MoId = -1, int OspDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "read");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetBarcode(MoId, OspDetailId);
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

        #region //GetOspBarcode 取得託外生產條碼
        [HttpPost]
        public void GetOspBarcode(int OspBarcodeId = -1, int OspDetailId = -1, string BarcodeNo = "", int MoProcessId = -1, int MoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "read");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspBarcode(OspBarcodeId, OspDetailId, BarcodeNo, MoProcessId, MoId
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

        #region //GetOspReceipt 取得託外入庫資料
        [HttpPost]
        public void GetOspReceipt(int OsrId = -1, string OsrErpFullNo = "", int SupplierId = -1, string OsrStartDate = "", string OsrEndDate = "", string OspNo = "", string WoErpFullNo = "", string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "read,constrained-data");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspReceipt(OsrId, OsrErpFullNo, SupplierId, OsrStartDate, OsrEndDate, OspNo, WoErpFullNo, MtlItemNo
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

        #region //GetOspReceiptDetail 取得託外入庫詳細資料
        [HttpPost]
        public void GetOspReceiptDetail(int OsrId = -1, int OsrDetailId = -1, string OspNo = "", string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = "", int InventoryId = -1, string OsrIdList =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "read,constrained-data");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspReceiptDetail(OsrId, OsrDetailId, OspNo, WoErpFullNo, MtlItemNo, MtlItemName, InventoryId, OsrIdList
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

        #region //GetOspReceiptBarcode 取得託外入庫條碼
        [HttpPost]
        public void GetOspReceiptBarcode(int OsrDetailId = -1, int OspBarcodeId = -1, string BarcodeNo = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "read");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspReceiptBarcode(OsrDetailId, OspBarcodeId, BarcodeNo, Status
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

        #region //GetSupplierMachine 取得供應商對應機台
        [HttpPost]
        public void GetSupplierMachine(int SupplierId = -1, int ProcessId = -1, int MachineId = -1)
        {
            try
            {
                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetSupplierMachine(SupplierId, ProcessId , MachineId);
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

        #region //GetOspDetail2 取得託外生產單詳細資料
        [HttpPost]
        public void GetOspDetail2(int OspDetailId = -1, int OspId = -1, int MoId = -1, string Status = "", int SupplierId = -1, string WoErpFullNo = "", string OspNo = "", string ProcessAlias = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspDetail(OspDetailId, OspId, MoId, Status, SupplierId, WoErpFullNo, OspNo
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

        #region //GetBarcodeExcel 取得製令條碼資料(Excel用)
        [HttpPost]
        public void GetBarcodeExcel(int OspId = -1, string MoFullNo = "", string MoProcess = "", int OspDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "read");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetBarcodeExcel(OspId, MoFullNo, MoProcess, OspDetailId);
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

        #region //托外加工商過站系統使用
        #region //廠商登入
        [HttpPost]
        public void ProductionHistorySupplierLogin(int SupplierCode)
        {
            try
            {
                string OspDetailDataRequest = "";
                JObject OspDetailResponse = new JObject();

                string MachineDataRequest = "";
                JObject MachineResponse = new JObject();

                string OspUserDataRequest = "";
                JObject OspUserResponse = new JObject();

                int MoId = -1;
                int MoProcessId = -1;
                int MachineId = 2810;
                int UserId = 5369;
                string Account = "";
                string Password = "";
                #region //Request
                OspDetailDataRequest = outsourcingProductionDA.GetOspDetail(SupplierCode, -1, -1, "", -1, "", "", "", -1, -1);
                OspDetailResponse = BaseHelper.DAResponse(OspDetailDataRequest);
                if (OspDetailResponse["status"].ToString() == "success")
                {
                    if(OspDetailResponse["result"].ToString() =="[]") throw new SystemException("廠商碼找不到資料，請重新確認!!");
                    var OspDetailDataResult = JObject.Parse(OspDetailDataRequest)["data"];
                    MoId = Convert.ToInt32(OspDetailDataResult[0]["MoId"].ToString());
                    MoProcessId = Convert.ToInt32(OspDetailDataResult[0]["MoProcessId"].ToString());

                    //MachineDataRequest = outsourcingProductionDA.GetOspDetailMachine(MoId, MoProcessId);
                    //MachineResponse = BaseHelper.DAResponse(MachineDataRequest);
                    //if (MachineResponse["status"].ToString() == "success")
                    //{
                    //    var MachineDataResult = JObject.Parse(MachineDataRequest)["data"];
                    //    MachineId = Convert.ToInt32(MachineDataResult[0]["MachineId"].ToString());

                    //}

                    OspUserDataRequest = outsourcingProductionDA.GetOspUser(UserId, "");
                    OspUserResponse = BaseHelper.DAResponse(OspUserDataRequest);
                    if (OspUserResponse["status"].ToString() == "success")
                    {
                        var OspUserDataResult = JObject.Parse(OspUserDataRequest)["data"];
                        Account = OspUserDataResult[0]["UserNo"].ToString();

                    }
                }
                #endregion

                #region //Response
                string dataRequest = JObject.FromObject(new
                {
                    status = "success",
                    data = new
                    {
                        MoId,
                        MoProcessId,
                        MachineId,
                        UserId,
                        Account
                    }
                }).ToString();

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

        #region //GetOspDetailData 托外過站介面-取得托外單身相關資料
        [HttpPost]
        public void GetOspDetailData(int OspDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("OspReceipt", "read");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspDetailData(OspDetailId);
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
        
        #region //GetOspMoProcess 取得託外加工製程(過站下拉用)
        [HttpPost]
        public void GetOspMoProcess(int OspDetailId = -1)
        {
            try
            {
                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspMoProcess(OspDetailId);
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

        #region //UpdateOspNextMoProcessForLotMode 過站-更新條碼下一站製程(批量修改)
        [HttpPost]
        public void UpdateOspNextMoProcessForLotMode(int OspDetailId = -1, string BarcodeNoListString = "", int MoProcessId = -1, int NewMoProcessId = -1)
        {
            try
            {
                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspNextMoProcessForLotMode(OspDetailId, BarcodeNoListString, MoProcessId, NewMoProcessId);
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

        #region //採用ProductionHistoryDA
        #region //TxBarcodeProcess -- BARCODE DATA COLLECT(目前過站主程式)
        [HttpPost]
        public void TxBarcodeProcess(string BarcodeNo = "", int MoProcessId = -1, int MachineId = -1, int UserId = -1,
            string QtyBarcodeStatus = "", string QtyLotNgCauseNo = "", int QtyNextMoProcessId = -1, string Company = "", int CompanyId = -1, string UserNo = "")
        {
            try
            {

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                var a = Session["CompanyId"];
                dataRequest = productionHistoryDA.TxBarcodeProcess(BarcodeNo, MoProcessId, MachineId, UserId
                    , QtyBarcodeStatus, QtyLotNgCauseNo, QtyNextMoProcessId, Company, UserNo);
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

        #region //GetBarcodeMoprocess 取得模仁條碼製程資料
        [HttpPost]
        public void GetBarcodeMoprocess(int NextMoProcessId = -1)
        {
            try
            {
                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeMoprocess(NextMoProcessId);
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

        #region //GetProcessQty 取得待加工數
        [HttpPost]
        public void GetProcessQty(int MoProcessId = -1)
        {
            try
            {
                #region //Request   
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetProcessQty(MoProcessId);
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

        #region //GetBarcodeProcess 取得製令條碼歷程
        [HttpPost]
        public void GetBarcodeProcess(int BarcodeProcessId = -1, int MoId = -1, int BarcodeId = -1, string BarcodeNo = "", int MoProcessId = -1)
        {
            try
            {

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeProcess(BarcodeProcessId, MoId, BarcodeId, BarcodeNo, MoProcessId);
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

        #region //GetActiveBarcode 取得尚未完工條碼
        [HttpPost]
        public void GetActiveBarcode(int MoProcessId = -1)
        {
            try
            {
                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetActiveBarcode(MoProcessId);
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
        #endregion

        #endregion

        #region //Add
        #region //AddOutsourcingProduction 新增託外生產單資料
        [HttpPost]
        public void AddOutsourcingProduction(int DepartmentId = -1, string OspNo = "", string OspDate = "", int SupplierId = -1, string OspStatus = ""
            , string OspDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOutsourcingProduction(DepartmentId, OspNo, OspDate, SupplierId, OspStatus
                    , OspDesc, Status);
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

        #region //AddOspDetail 新增託外生產單詳細資料
        [HttpPost]
        public void AddOspDetail(int OspId = -1, int MoId = -1, int MoProcessId = -1, string ProcessCheckStatus = "", string ProcessCheckType = "", int OspQty = -1, int SuppliedQty = -1
            , string ProcessCode = "", string ExpectedDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOspDetail(OspId, MoId, MoProcessId, ProcessCheckStatus, ProcessCheckType, OspQty, SuppliedQty
                    , ProcessCode, ExpectedDate, Remark);
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

        #region //AddOspBarcode 新增託外生產條碼
        [HttpPost]
        public void AddOspBarcode(int OspDetailId = -1, string BarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOspBarcode(OspDetailId, BarcodeNo, MoProcessId);
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

        #region //AddOspReceipt 新增託外入庫資料
        [HttpPost]
        public void AddOspReceipt(int SupplierId = -1, string SupplierSo = "", string OsrErpPrefix = "", string OsrErpNo = "", string DocumentDate = "", string ReceiptDate = "", string ReserveTaxCode = "", string TaxCode = ""
            , string TaxType = "", string InvoiceType = "", double TaxRate = -1, string UiNo = "", string InvoiceDate = "", string InvoiceNo = "", string ApplyYYMM = "", string CurrencyCode = "", double Exchange = -1, int RowCnt = -1
            , string Remark = "", string DeductType = "", string PaymentTerm = "")
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOspReceipt(SupplierId, SupplierSo, OsrErpPrefix, OsrErpNo, DocumentDate, ReceiptDate, ReserveTaxCode, TaxCode
                    , TaxType, InvoiceType, TaxRate, UiNo, InvoiceDate, InvoiceNo, ApplyYYMM, CurrencyCode, Exchange, RowCnt
                    , Remark, DeductType, PaymentTerm);
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

        #region //AddOspReceiptDetail 新增託外入庫詳細資料
        [HttpPost]
        public void AddOspReceiptDetail(int OsrId = -1, int OspDetailId = -1, string OsrSeq = "", string LotNumber = "",int InventoryId = -1, string AcceptanceDate = ""
            , int ReceiptQty = -1, string AvailableUom = "", int AcceptQty = -1, int ScriptQty = -1, int ReturnQty = -1, int AvailableQty = -1, double OrigUnitPrice = -1, double OrigAmount = -1
            , string DiscountDescription = "", string Overdue = "", string Remark = "", double OrigPreTaxAmt = -1, double OrigTaxAmt = -1, double OrigDiscountAmt = -1, double PreTaxAmt = -1, double TaxAmt = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOspReceiptDetail(OsrId, OspDetailId, OsrSeq, LotNumber, InventoryId, AcceptanceDate
                    , ReceiptQty, AvailableUom, AcceptQty, ScriptQty, ReturnQty, AvailableQty, OrigUnitPrice, OrigAmount
                    , DiscountDescription, Overdue, Remark, OrigPreTaxAmt, OrigTaxAmt, OrigDiscountAmt, PreTaxAmt, TaxAmt);
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

        #region //AddOspReceiptBarcode 新增託外入庫條碼資料
        [HttpPost]
        public void AddOspReceiptBarcode(string OsrReceiptBarcodeData = "")
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOspReceiptBarcode(OsrReceiptBarcodeData);
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

        #region //BatchAddOspReceiptDetail 批量新增託外入庫資料
        [HttpPost]
        public void BatchAddOspReceiptDetail(int OsrId = -1, string OspDetailIdList = "")
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "add");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.BatchAddOspReceiptDetail(OsrId, OspDetailIdList);
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

        #region //AddOspDetailExcel 新增託外生產單詳細資料(批量)
        [HttpPost]
        public void AddOspDetailExcel(int OspId = -1, string ExcelData = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "batch");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.AddOspDetailExcel(OspId, ExcelData);
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

        #region //Update
        #region //UpdateOutsourcingProduction 更新託外生產單資料
        [HttpPost]
        public void UpdateOutsourcingProduction(int OspId = -1, int DepartmentId = -1, string OspNo = "", string OspDate = "", int SupplierId = -1, string OspStatus = ""
            , string OspDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOutsourcingProduction(OspId, DepartmentId, OspNo, OspDate, SupplierId, OspStatus
                    , OspDesc, Status);
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

        #region //UpdateOspDetail 更新託外生產單詳細資料
        [HttpPost]
        public void UpdateOspDetail(int OspDetailId = -1, int OspId = -1, int MoId = -1, int MoProcessId = -1, string ProcessCheckStatus = "", string ProcessCheckType = "", int OspQty = -1, int SuppliedQty = -1
            , string ProcessCode = "", string ExpectedDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspDetail(OspDetailId, OspId, MoId, MoProcessId, ProcessCheckStatus, ProcessCheckType, OspQty, SuppliedQty
                    , ProcessCode, ExpectedDate, Remark);
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

        #region //UpdateOspConfirm 確認託外生產單資料
        [HttpPost]
        public void UpdateOspConfirm(int OspId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspConfirm(OspId);
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

        #region //UpdateOspReConfirm 反確認託外生產單資料
        [HttpPost]
        public void UpdateOspReConfirm(int OspId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspReConfirm(OspId);
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

        #region //UpdateOspReceipt 更新託外入庫單資料
        [HttpPost]
        public void UpdateOspReceipt(int OsrId = -1, int SupplierId = -1, string SupplierSo = "", string OsrErpPrefix = "", string OsrErpNo = "", string DocumentDate = "", string ReceiptDate = "", string ReserveTaxCode = "", string TaxCode = ""
            , string TaxType = "", string InvoiceType = "", double TaxRate = -1, string UiNo = "", string InvoiceDate = "", string InvoiceNo = "", string ApplyYYMM = "", string CurrencyCode = "", double Exchange = -1, int RowCnt = -1
            , string Remark = "", string DeductType = "", string PaymentTerm = "")
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspReceipt(OsrId, SupplierId, SupplierSo, OsrErpPrefix, OsrErpNo, DocumentDate, ReceiptDate, ReserveTaxCode, TaxCode
                    , TaxType, InvoiceType, TaxRate, UiNo, InvoiceDate, InvoiceNo, ApplyYYMM, CurrencyCode, Exchange, RowCnt
                    , Remark, DeductType, PaymentTerm);
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

        #region //UpdateOspReceiptDetail 更新託外入庫單詳細資料
        [HttpPost]
        public void UpdateOspReceiptDetail(int OsrDetailId = -1, int OsrId = -1, int OspDetailId = -1, string OsrSeq = "", string LotNumber = "", int InventoryId = -1, string AcceptanceDate = ""
            , int ReceiptQty = -1, string AvailableUom = "", int AcceptQty = -1, int ScriptQty = -1, int ReturnQty = -1, int AvailableQty = -1, double OrigUnitPrice = -1, double OrigAmount = -1
            , string DiscountDescription = "", string Overdue = "", string Remark = "", double OrigPreTaxAmt = -1, double OrigTaxAmt = -1, double OrigDiscountAmt = -1, double PreTaxAmt = -1, double TaxAmt = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspReceiptDetail(OsrDetailId, OsrId, OspDetailId, OsrSeq, LotNumber, InventoryId, AcceptanceDate
                    , ReceiptQty, AvailableUom, AcceptQty, ScriptQty, ReturnQty, AvailableQty, OrigUnitPrice, OrigAmount
                    , DiscountDescription, Overdue, Remark, OrigPreTaxAmt, OrigTaxAmt, OrigDiscountAmt, PreTaxAmt, TaxAmt);
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

        #region //UpdateTransferOspReceipt 拋轉託外入庫單資料(不核單)
        [HttpPost]
        public void UpdateTransferOspReceipt(int OsrId)
        {
            try
            {                
                WebApiLoginCheck("OspReceipt", "confirm");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateTransferOspReceipt(OsrId);
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

        #region //UpdateConfirmOspReceipt 確認託外入庫單資料
        [HttpPost]
        public void UpdateConfirmOspReceipt(int OsrId)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "confirm");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateConfirmOspReceipt(OsrId);
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

        #region //UpdateTACOspReceipt 拋轉並核准託外入庫單資料
        [HttpPost]
        public void UpdateTACOspReceipt(int OsrId)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "confirm");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateTACOspReceipt(OsrId);
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

        #region //UpdateReConfirmOspReceipt 反確認託外入庫單資料
        [HttpPost]
        public void UpdateReConfirmOspReceipt(int OsrId)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "reconfirm");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateReConfirmOspReceipt(OsrId);
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

        #region //UpdateVoidOspReceipt 作廢託外入庫單資料
        [HttpPost]
        public void UpdateVoidOspReceipt(int OsrId)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "void");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateVoidOspReceipt(OsrId);
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

        #region //UpdateBarcodeProcess 託外入庫條碼過站
        [HttpPost]
        public void UpdateBarcodeProcess(int OsrDetailId)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "barcode-process");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateBarcodeProcess(OsrDetailId);
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

        #region //UpdateOsrQuantity 更新託外進貨單頭數量
        [HttpPost]
        public void UpdateOsrQuantity(int OsrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOsrQuantity(OsrDetailId);
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

        #region //UpdateOsrDetailScriptQty 更新託外入庫單身報廢數量
        [HttpPost]
        public void UpdateOsrDetailScriptQty(int OsrDetailId = -1, double ScriptQty = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOsrDetailScriptQty(OsrDetailId, ScriptQty);
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

        #region //UpdateOspVoid 託外生產單作廢
        [HttpPost]
        public void UpdateOspVoid(int OspId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "update");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspVoid(OspId);
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

        #region //UpdateOspConfirmBatch 確認託外生產單資料(批量)
        [HttpPost]
        public void UpdateOspConfirmBatch(string OspIdList = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "batch");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspConfirmBatch(OspIdList);
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

        #region //UpdateOspReConfirmBatch 反確認託外生產單資料(批量)
        [HttpPost]
        public void UpdateOspReConfirmBatch(string OspIdList = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "batch");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.UpdateOspReConfirmBatch(OspIdList);
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

        #region //Delete
        #region //DeleteOutsourcingProduction -- 刪除託外生產單資料
        [HttpPost]
        public void DeleteOutsourcingProduction(int OspId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "delete");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOutsourcingProduction(OspId);
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

        #region //DeleteOspBarcode -- 刪除託外生產條碼
        [HttpPost]
        public void DeleteOspBarcode(string OspBarcodeId = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "delete");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspBarcode(OspBarcodeId);
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

        #region //DeleteOspDetail -- 刪除託外生產詳細資料
        [HttpPost]
        public void DeleteOspDetail(int OspDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "delete");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspDetail(OspDetailId);
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

        #region //DeleteOspReceipt -- 刪除託外入庫資料
        [HttpPost]
        public void DeleteOspReceipt(int OsrId = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "delete");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspReceipt(OsrId);
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

        #region //DeleteOspReceiptDetail -- 刪除託外入庫詳細資料
        [HttpPost]
        public void DeleteOspReceiptDetail(int OsrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "delete");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspReceiptDetail(OsrDetailId);
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

        #region //DeleteOspReceiptBarcode -- 刪除託外入庫條碼資料
        [HttpPost]
        public void DeleteOspReceiptBarcode(int OsrDetailId = -1, int OspBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "delete");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspReceiptBarcode(OsrDetailId, OspBarcodeId);
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

        #region //DeleteOspBatch 刪除託外生產單資料(批量)
        [HttpPost]
        public void DeleteOspBatch(string OspIdList = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "batch");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspBatch(OspIdList);
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

        #region //DeleteOspDetailExcel -- 刪除託外生產詳細資料(批量)
        [HttpPost]
        public void DeleteOspDetailExcel(string ExcelData = "")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "batch");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.DeleteOspDetailExcel(ExcelData);
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

        #region //BPM
        #region //TransferOspToBpm 拋轉托外生產單至BPM
        [HttpPost]
        public void TransferOspToBpm(string OspIds)
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "transfer");

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.TransferOspToBpm(OspIds);
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

        #region //GetBpmOspStatusData -- 取得BPM委外單回傳狀態資料
        [HttpPost]
        [Route("api/bpm/updateOspStatus")]
        public void GetBpmOspStatusData(string id = "", string status = "", string rootId = "", string comfirmUser = "", string bpmNo = "")
        {
            try
            {
                var dataRequestJson = new JObject();

                //先記錄LOG
                dataRequest = outsourcingProductionDA.AddOspBpmLog(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser);
                dataRequestJson = JObject.Parse(dataRequest);
                int LogId = -1;
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    foreach (var item in dataRequestJson["data"])
                    {
                        LogId = Convert.ToInt32(item["LogId"]);
                    }
                }

                if (status != "Y" && status != "E")
                {
                    throw new SystemException("回傳狀態錯誤!!");
                }

                #region //更改MES請購單狀態
                dataRequest = outsourcingProductionDA.UpdateOspStatus(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser);
                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
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
        #endregion

        #region//Download
        #region //Pdf

        #region//託外加工單(MES)
        public JObject HeadjsonResponse = new JObject();
        public JObject LinejsonResponse = new JObject();
        public JObject BarcodejsonResponse = new JObject();
        public string HeadDataRequest = "";
        public string LineDataRequest = "";
        public string BarcodeDataRequest = "";
        public void GetOspCardPdf(int OspId = -1,string BarcodeStatus ="")
        {
            try
            {
                WebApiLoginCheck("OutsourcingProduction", "print");
                outsourcingProductionDA = new OutsourcingProductionDA();

                #region//產生PDF
                WebClient wc = new WebClient();
                string htmlText = "", htmlDetail = "", htmlDetail2 = "", OspNo = "";
                string PassStationControl = "";
                string TableLineAll = "";
                string tableLine = "";

                List<string> ospDetailList = new List<string>();

                #region //Request 單頭
                HeadDataRequest = outsourcingProductionDA.GetOutsourcingProduction(OspId, -1, "", -1, "", "", "", "", "", "", "", -1, -1);
                HeadjsonResponse = BaseHelper.DAResponse(HeadDataRequest);
                #endregion

                #region //Request 單身
                LineDataRequest = outsourcingProductionDA.GetOspDetail(-1, OspId, -1, "", -1, "", "", "", -1, -1);
                LinejsonResponse = BaseHelper.DAResponse(LineDataRequest);
                #endregion

                

                int row = int.Parse(LinejsonResponse["result"][0]["TotalCount"].ToString());
                int page = 4;
                int pageNum = 0;
                if (row % 4 != 0)
                {
                    pageNum = (row / page) + 1;
                }
                else
                {
                    pageNum = row / page;
                }
                int j = 0;
                int i = 0;
                for (i = 0; i < 1; i++)
                {

                    if (i == 0)
                    {
                        j = 0;
                    }
                    else
                    {
                        j = j * i;
                    }

                    if (HeadjsonResponse["status"].ToString() == "success")
                    {
                        var HeadDataResult = JObject.Parse(HeadDataRequest)["data"];
                        OspNo = HeadDataResult[0]["OspNo"].ToString();
                        PassStationControl = HeadDataResult[0]["PassStationControl"].ToString();    

                        #region //html
                        htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/OspPdf2.0.html"));
                        string CompanyName = "", CompanyAddress = "";
                        if (HeadDataResult[0]["CompanyId"].ToString() == "2")
                        {
                            CompanyName = "中揚光電股份有限公司";
                            CompanyAddress = HeadDataResult[0]["Address"].ToString();
                        }
                        else if (HeadDataResult[0]["CompanyId"].ToString() == "4")
                        {
                            CompanyName = "晶彩光學有限公司";
                            CompanyAddress = HeadDataResult[0]["Address"].ToString();
                        }
                        else if (HeadDataResult[0]["CompanyId"].ToString() == "5")
                        {
                            CompanyName = "群英光學有限公司";
                            CompanyAddress = HeadDataResult[0]["Address"].ToString();
                        }
                        else if (HeadDataResult[0]["CompanyId"].ToString() == "3")
                        {
                            CompanyName = "紘立光電有限公司";
                            CompanyAddress = HeadDataResult[0]["Address"].ToString();
                        }

                        htmlText = htmlText.Replace("[CompanyName]", CompanyName);
                        htmlText = htmlText.Replace("[CompanyAddress]", CompanyAddress);
                        htmlText = htmlText.Replace("[OspNo]", HeadDataResult[0]["OspNo"].ToString());
                        htmlText = htmlText.Replace("[SupplierShortName]", HeadDataResult[0]["SupplierShortName"].ToString());
                        htmlText = htmlText.Replace("[TelNoFirst]", HeadDataResult[0]["TelNoFirst"].ToString());
                        htmlText = htmlText.Replace("[FaxNo]", HeadDataResult[0]["FaxNo"].ToString());

                        htmlText = htmlText.Replace("[OspDate]", HeadDataResult[0]["OspDate"].ToString());
                        htmlText = htmlText.Replace("[SupplierNo]", HeadDataResult[0]["SupplierNo"].ToString());

                        htmlText = htmlText.Replace("[OspStatus]", HeadDataResult[0]["OspStatus"].ToString());
                        htmlText = htmlText.Replace("[DepartmentName]", HeadDataResult[0]["DepartmentName"].ToString());

                        if (HeadDataResult[0]["OspDesc"].ToString() == "")
                        {
                            htmlText = htmlText.Replace("[OspDesc]", "備註:");
                        }
                        else
                        {
                            htmlText = htmlText.Replace("[OspDesc]", "備註:" + HeadDataResult[0]["OspDesc"].ToString());
                        }
                        string htmlTemplate = htmlText;
                        #endregion

                    }
                    if (LinejsonResponse["status"].ToString() == "success")
                    {
                        //取出在OspId底下的OspDetailId
                        dynamic[] result = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(LineDataRequest)["data"].ToString());
                        if (page + (page * i) > row)
                        {
                            page = row;
                        }
                        else
                        {
                            page = page + (page * i);
                        }
                        //for (int x = j; j < page; j++)
                        for (j = 0; j < row; j++)
                        {
                            htmlDetail += @"<tr>
                                <th rowspan='2' style='text-align: center;'>" + (j + 1);
                            htmlDetail += @"</th>
                                <td style='text-align: center; height:3%'>[MtlItemNo]</td>
                                <td style='text-align: center; height:3%'><b>[WoErpPrefix]-[WoErpNo]([WoSeq])</b></td>
                                <td rowspan='2' style='text-align: center; height:3%'>[MtlItemSpec]</td>
                                <td rowspan='2' style='text-align: center; height:3%'>[ExpectedDate]</td>
                                <td rowspan='2' style='text-align: center; height:3%;width:7%'>[RoutingItemProcessDesc]</td>
                                <td rowspan='2' style='text-align: center; height:3%'>[ProcessCheckStatus]</td>
                                <td rowspan='2' style='text-align: center; height:3%'>[OspQty]</td>
                                <td rowspan='2' style='text-align: center; height:3%'>[SuppliedQty]</td>
                                <td rowspan='2'></td>
                                <td rowspan='2'></td>
                                <td rowspan='2'></td>
                                <td rowspan='2'></td>
                                <td rowspan='2'></td>
                                <td rowspan='2'></td>
                            </tr>
                            <tr>
                                <td style='text-align: center;'>[MtlItemName]</td>
                                <td style='text-align: center;'>[ProcessCodeName]([ProcessCode])</td>
                            </tr>";
                            htmlDetail = htmlDetail.Replace("[MtlItemNo]", Convert.ToString(result[j]["MtlItemNo"]));
                            htmlDetail = htmlDetail.Replace("[WoErpPrefix]", Convert.ToString(result[j]["WoErpPrefix"]));
                            htmlDetail = htmlDetail.Replace("[WoErpNo]", Convert.ToString(result[j]["WoErpNo"]));
                            htmlDetail = htmlDetail.Replace("[WoSeq]", Convert.ToString(result[j]["WoSeq"]));
                            htmlDetail = htmlDetail.Replace("[MtlItemSpec]", Convert.ToString(result[j]["MtlItemSpec"]));                            
                            htmlDetail = htmlDetail.Replace("[ExpectedDate]", Convert.ToString(result[j]["ExpectedDate"]));
                            htmlDetail = htmlDetail.Replace("[RoutingItemProcessDesc]", Convert.ToString(result[j]["RoutingItemProcessDesc"]));
                            htmlDetail = htmlDetail.Replace("[ProcessCheckStatus]", Convert.ToString(result[j]["ProcessCheckStatus"]));
                            htmlDetail = htmlDetail.Replace("[OspQty]", Convert.ToString(result[j]["OspQty"]));
                            htmlDetail = htmlDetail.Replace("[SuppliedQty]", Convert.ToString(result[j]["SuppliedQty"]));
                            htmlDetail = htmlDetail.Replace("[MtlItemName]", Convert.ToString(HttpUtility.HtmlEncode(result[j]["MtlItemName"])));
                            htmlDetail = htmlDetail.Replace("[ProcessCodeName]", Convert.ToString(result[j]["ProcessCodeName"]));
                            htmlDetail = htmlDetail.Replace("[ProcessCode]", Convert.ToString(result[j]["ProcessCode"]));
                        }   
                    }                    
                }
                htmlText = htmlText.Replace("[OspDetail]", htmlDetail);
                if (PassStationControl == "Y")
                {
                    if (LinejsonResponse["status"].ToString() == "success")
                    {
                        //取出在OspId底下的OspDetailId
                        dynamic[] result = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(LineDataRequest)["data"].ToString());
                        page = 4;

                        if (page + (page * i) > row)
                        {
                            page = row;
                        }
                        else
                        {
                            page = page + (page * i);
                        }
                        //for (int x = j; j < page; j++)
                        
                        for (j = 0; j < row; j++)
                        {
                            #region 當非條碼版本控制Line每三個就強制分頁
                            string PageSetting = "";
                            if(BarcodeStatus != "N")
                            {
                                PageSetting = "avoid";
                            }
                            else
                            {
                                if ((j + 1) % 3 == 0)
                                {
                                    PageSetting = "always";
                                }
                                else
                                {
                                    PageSetting = "auto";
                                }
                            }
                            #endregion

                            #region tableLine初始化
                            tableLine = @"
                                    <table class='tableLine' style='width:100%; margin-bottom:0;'>
                                        <thead>
                                            <tr>
                                                <th style='text-align: center; width:12.5%;' rowspan='2'>項次</th>
                                                <th style='text-align: center; width:12.5%;'>料號(型號)</th>
                                                <th style='text-align: center; width:12.5%;'>規格</th>
                                                <th style='text-align: center; width:12.5%;'>製程</th>
                                                <th style='text-align: center; width:12.5%;' rowspan='2'>預計交期</th>
                                                <th style='text-align: center; width:12.5%;' rowspan='2'>加工細項</th>
                                                <th style='text-align: center; width:12.5%;' rowspan='2'>是否QA檢</th>
                                                <th style='text-align: center; width:12.5%;' rowspan='2'>過站登入條碼</th>
                                            </tr>
                                            <tr>
                                                <th style='text-align: center;'>品名</th>
                                                <th style='text-align: center;'>製令單號</th>
                                                <th style='text-align: center;'>加工代碼</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            [LineData]
                                        </tbody>
                                    </table>
                            ";
                            #endregion

                            #region LineData導入
                            #region LineData資料建立
                            htmlDetail2 = "";
                            htmlDetail2 += @"<tr>
                                            <th rowspan='2' style='text-align: center;'>" + (j + 1);
                            htmlDetail2 += @"</th>
                                <td style='text-align: center;'>[MtlItemNo]</td>
                                <td style='text-align: center;'>[MtlItemSpec]</td>
                                <td style='text-align: center; height:3%;'>[ProcessAlias]</td>
                                <td rowspan='2' style='text-align: center; height:3%;'>[ExpectedDate]</td>
                                <td rowspan='2' style='text-align: center; height:3%;'>[RoutingItemProcessDesc]</td>
                                <td rowspan='2' style='text-align: center; height:3%;'>[ProcessCheckStatus]</td>
                                <td style='text-align: center; height:50px; border:0px;'>
                                    <img style='width:38px' src='[srcOspDetailId]'/>
                                </td>
                            </tr>
                            <tr>
                                <td style='text-align: center;'>[MtlItemName]</td>
                                <td style='text-align: center;'><b>[WoErpPrefix]-[WoErpNo]([WoSeq])</b></td>
                                <td style='text-align: center;'>[ProcessCodeName]([ProcessCode])</td>
                                <td style='text-align: center; border:0px;page-break-after: " + PageSetting+@";'>[OspDetailId]</td>
                            </tr>";
                            if (BarcodeStatus != "N")
                            {
                                htmlDetail2 += @"<tr>
                                                    <th colspan='1' style='text-align:center;height:35px;'>條碼列表</th>
                                                    <th colspan='7' style='text-align:center;height:50px;'>&nbsp;</th>
                                                </tr>
                                                [BarcodeList]
                                                ";
                            }
                            htmlDetail2 = htmlDetail2.Replace("[MtlItemNo]", Convert.ToString(result[j]["MtlItemNo"]));
                            htmlDetail2 = htmlDetail2.Replace("[WoErpPrefix]", Convert.ToString(result[j]["WoErpPrefix"]));
                            htmlDetail2 = htmlDetail2.Replace("[WoErpNo]", Convert.ToString(result[j]["WoErpNo"]));
                            htmlDetail2 = htmlDetail2.Replace("[WoSeq]", Convert.ToString(result[j]["WoSeq"]));
                            htmlDetail2 = htmlDetail2.Replace("[MtlItemSpec]", Convert.ToString(result[j]["MtlItemSpec"]));
                            htmlDetail2 = htmlDetail2.Replace("[ExpectedDate]", Convert.ToString(result[j]["ExpectedDate"]));
                            htmlDetail2 = htmlDetail2.Replace("[ProcessAlias]", Convert.ToString(result[j]["ProcessAlias"]));
                            htmlDetail2 = htmlDetail2.Replace("[RoutingItemProcessDesc]", Convert.ToString(result[j]["RoutingItemProcessDesc"]));
                            htmlDetail2 = htmlDetail2.Replace("[ProcessCheckStatus]", Convert.ToString(result[j]["ProcessCheckStatus"]));
                            htmlDetail2 = htmlDetail2.Replace("[MtlItemName]", Convert.ToString(result[j]["MtlItemName"]));
                            htmlDetail2 = htmlDetail2.Replace("[ProcessCodeName]", Convert.ToString(result[j]["ProcessCodeName"]));
                            htmlDetail2 = htmlDetail2.Replace("[ProcessCode]", Convert.ToString(result[j]["ProcessCode"]));

                            
                            int OspDetailId = Convert.ToInt32(result[j]["OspDetailId"]);
                            int MoProcessId = Convert.ToInt32(result[j]["MoProcessId"]);
                            int MoId = Convert.ToInt32(result[j]["MoId"]);
                            

                            #region QRcode樣式設定
                            var writer = new BarcodeWriter  //dll裡面可以看到屬性
                            {
                                Format = BarcodeFormat.QR_CODE,
                                Options = new EncodingOptions //設定大小
                                {
                                    Height = 10,
                                    Width = 10,
                                    PureBarcode = true,
                                    Margin = 0
                                }
                            };
                            #endregion

                            #region 產生OspDetailId QRcode
                            var img = writer.Write(Convert.ToString(OspDetailId));
                            Bitmap myBitmap = new Bitmap(img);
                            string FileName = "OspDetailId" + OspDetailId;
                            string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/OSP/OspDetailId/{0}.bmp", FileName));
                            myBitmap.Save(filePath, ImageFormat.Bmp);
                            //帶入OspDetailId圖片
                            htmlDetail2 = htmlDetail2.Replace("[srcOspDetailId]", Server.MapPath("~/PdfTemplate/MES/OSP/OspDetailId/OspDetailId" + OspDetailId + ".bmp"));
                            htmlDetail2 = htmlDetail2.Replace("[OspDetailId]", OspDetailId.ToString());
                            #endregion
                            #endregion

                            #region //條碼列表建立
                            if (BarcodeStatus != "N")
                            {
                                string BarcodeListAll = "";
                                string BarcodeList = "";
                                string BarcodeNoListAll = "";
                                string BarcodeNoList = "";

                                #region 取得OspDetailId的條碼資料
                                BarcodeDataRequest = outsourcingProductionDA.GetOspBarcode(-1, OspDetailId, "", MoProcessId, MoId, "", -1, -1);
                                BarcodejsonResponse = BaseHelper.DAResponse(BarcodeDataRequest);
                                #endregion

                                if (BarcodejsonResponse["status"].ToString() == "success")
                                {
                                    dynamic[] resultBarcode = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(BarcodeDataRequest)["data"].ToString());
                                    int num = resultBarcode.Count();
                                    int nowNum = 0;
                                    for (var k = 0; k < num; k++)
                                    {
                                        nowNum++;
                                        int OspBarcodeId = Convert.ToInt32(resultBarcode[k]["BarcodeId"]);
                                        string OspBarcodeNo = Convert.ToString(resultBarcode[k]["BarcodeNo"]);

                                        #region QRcode樣式設定
                                    writer = new BarcodeWriter  //dll裡面可以看到屬性
                                    {
                                        Format = BarcodeFormat.QR_CODE,
                                        Options = new EncodingOptions //設定大小
                                        {
                                            Height = 10,
                                            Width = 10,
                                            PureBarcode = true,
                                            Margin = 0
                                        }
                                    };
                                    #endregion

                                        #region BarcodeId 產生QRcode
                                        img = writer.Write(Convert.ToString(OspBarcodeNo));
                                        myBitmap = new Bitmap(img);
                                        FileName = "OspBarcode" + OspBarcodeId;
                                        filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/OSP/OspBarcode/{0}.bmp", FileName));
                                        myBitmap.Save(filePath, ImageFormat.Bmp);
                                        #endregion

                                        #region //初始先加入<tr>
                                        if (nowNum % 8 == 1)
                                        {
                                            
                                            BarcodeList += "<tr>";
                                            BarcodeNoList += "<tr>";
                                        }
                                        #endregion

                                        #region //條碼列表圖片html結構
                                    BarcodeList += @"
                                        <td style='text-align:center; height:150px;' rowspan='3'>
                                            <table>
                                                <tr><td style='text-align:center;border:0; height:100px;'><img style='width:50px' src='[OspBarcode]'/></td></tr>
                                                <tr><td style='text-align:center;border:0;'>[BarcodeNo]</td></tr>
                                            </table>
                                        </td>
                                        ";
                                    #endregion

                                        #region //條碼列表文字html結構
                                    BarcodeNoList += @"
                                        <td style='text-align: center;'>
                                            [BarcodeNo]
                                        </td>
                                        ";
                                    #endregion

                                        #region //結構帶入
                                        BarcodeList = BarcodeList.Replace("[OspBarcode]", Server.MapPath("~/PdfTemplate/MES/OSP/OspBarcode/OspBarcode" + OspBarcodeId + ".bmp"));
                                        BarcodeList = BarcodeList.Replace("[BarcodeNo]", OspBarcodeNo.ToString());
                                        #endregion

                                        #region //第8個結束加入</tr>
                                        if (nowNum % 8 == 0 || k == num - 1)
                                        {
                                            if(k == num - 1)
                                            {
                                                #region //條碼列表圖片html結構
                                                if (nowNum % 8 != 0)
                                                {
                                                    var endNum = 0;
                                                    if (nowNum <= 8)
                                                    {
                                                        endNum = 8 - nowNum;
                                                    }
                                                    else
                                                    {
                                                        endNum = 8 - nowNum % 8;

                                                    }
                                                    for (int z = 1; z <= endNum; z++)
                                                    {
                                                        if (z == 8 - nowNum % 8)
                                                        {
                                                            BarcodeList += @"
                                                            <td rowspan='3' style='height:50px; page-break-after: always;'>
                                                                &nbsp;
                                                            </td>
                                                            ";
                                                        }
                                                        else
                                                        {
                                                            BarcodeList += @"
                                                            <td rowspan='3'>
                                                                &nbsp;
                                                            </td>
                                                            ";
                                                        }

                                                    }
                                                }
                                                #endregion
                                            }

                                            BarcodeList += "</tr>";
                                            BarcodeNoList += "</tr>";
                                            BarcodeListAll += BarcodeList;
                                            BarcodeNoListAll += BarcodeNoList;
                                            BarcodeList = "";
                                            BarcodeNoList = "";
                                        }
                                        #endregion
                                    }
                                    htmlDetail2 = htmlDetail2.Replace("[BarcodeList]", BarcodeListAll);
                                    //htmlDetail2 = htmlDetail2.Replace("[BarcodeNoList]", BarcodeNoListAll);
                                }
                            }
                            #endregion

                            //LineData導入
                            tableLine = tableLine.Replace("[LineData]", htmlDetail2);
                            //完整一個TableLine建立完畢
                            TableLineAll += tableLine;

                            #endregion

                        }
                        //TableLine導入
                        htmlText = htmlText.Replace("[TableLine]", TableLineAll);

                    }
                }
                else
                {
                    string pattern = @"<section\b[^>]*>(.*?)</section>";
                    htmlText = Regex.Replace(htmlText, pattern, "", RegexOptions.Singleline);
                }
                #endregion

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4.Rotate())) //要寫PDF的文件，建構子沒填的話預設直式A4(PageSize.A4);若要PDF為横向則使用PageSize.A4.Rotate()
                        {
                            PdfWriter pdfWriter = PdfWriter.GetInstance(document, output); //指定文件預設開檔時的縮放為100%
                            PdfDestination pdfDestination = new PdfDestination(PdfDestination.XYZ, 0, document.PageSize.Height, 1f); //開啟Document文件 

                            document.Open(); //開啟Document文件 

                            XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, document, input, null, Encoding.UTF8, new UnicodeFontFactory()); //使用XMLWorkerHelper把Html parse到PDF檔裡

                            PdfAction action = PdfAction.GotoLocalPage(1, pdfDestination, pdfWriter); //將pdfDest設定的資料寫到PDF檔

                            pdfWriter.SetOpenAction(action);
                        }
                    }

                    pdfFile = output.ToArray();
                }

                string fileGuid = Guid.NewGuid().ToString();
                Session[fileGuid] = pdfFile;
                #endregion

                #region //製令QR圖片刪除
                DirectoryInfo LinkOspDetailId = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/OSP/OspDetailId"));
                FileInfo[] OspFilesList = LinkOspDetailId.GetFiles();
                foreach (FileInfo file in OspFilesList)
                {
                    file.Delete();
                }
                #endregion

                #region //條碼圖片刪除
                DirectoryInfo LinkOspBarcode = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/OSP/OspBarcode"));
                FileInfo[] OspBarcodeList = LinkOspBarcode.GetFiles();
                foreach (FileInfo file in OspBarcodeList)
                {
                    file.Delete();
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = OspNo + "託外單資料",
                    fileExtension = ".pdf"
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

        #region //ExcelOspReceiptDetail 託外入庫彙整檔匯出Excel 
        public void ExcelOspReceiptDetail(int OsrId = -1, string OsrErpFullNo = "", int SupplierId = -1, string OsrStartDate = "", string OsrEndDate = "", string OspNo = "", string WoErpFullNo = "", string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck("OspReceipt", "excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.GetOspReceipt(OsrId, OsrErpFullNo, SupplierId, OsrStartDate, OsrEndDate, OspNo, WoErpFullNo, MtlItemNo
                    , "", -1, -1);
                dynamic[] dataOsrId = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                string OsrIdListStr = "";
                if(dataOsrId.Count()<=0) throw new SystemException("查無單據!無法執行匯出");
                foreach (var item in dataOsrId)
                {
                    OsrIdList.Add(item.OsrId.ToString());
                }
                #endregion

                #region //Request
                dataRequest = outsourcingProductionDA.GetOspReceiptDetail(-1, -1, "", "", "", "", -1, string.Join(", ", OsrIdList)
                    , "", -1, -1);
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
                    string excelFileName = "【MES2.0】託外入庫資料彙整檔";
                    string excelsheetName = "託外入庫資料彙整頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "ERP託外入庫單", "供應商", "開立人員", "拋轉狀態", "核單狀態", "過站狀態", "單據日期", "託外入庫日期", "序號", "託外生產單號", "製令", "品號", "品名", "規格", "單位", "進貨數量", "進貨庫別", "過站狀態", "託外製程", "加工代碼", "驗收數量", "計價數量", "報廢數量", "加工單價", "加工金額", "原幣扣款金額", "原幣未稅金額", "原幣稅金額", "本幣未稅金額", "本幣稅金額", "確認碼", "確認者" };
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
                        worksheet.Range(rowIndex, 1, rowIndex, 32).Merge().Style = titleStyle;
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

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.OsrFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.OsrSupplierName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.OsrUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.TransferStatus == "Y" ? "已拋轉" : "未拋轉";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.OsrConfirmStatus == "Y" ? "已核單" : "未核單";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.ProcessStatus == "Y" ? "已過站" : "未過站";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.DocumentDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ReceiptDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.OsrSeq.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.OspNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.WoErpPrefix.ToString() + "-" + item.WoErpNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.MtlItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.UomNo.ToString() +" "+ item.UomName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.ReceiptQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.InventoryNo.ToString() +" "+ item.InventoryName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.ProcessStatus == "Y"? "已過站": "未過站";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = item.ProcessAlias.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(20, rowIndex)).Value = item.ProcessCodeName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(21, rowIndex)).Value = item.AcceptQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(22, rowIndex)).Value = item.AvailableQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(23, rowIndex)).Value = item.ScriptQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(24, rowIndex)).Value = item.OrigUnitPrice.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(25, rowIndex)).Value = item.OrigAmount.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(26, rowIndex)).Value = item.OrigDiscountAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(27, rowIndex)).Value = item.OrigPreTaxAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(28, rowIndex)).Value = item.OrigTaxAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(29, rowIndex)).Value = item.PreTaxAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(30, rowIndex)).Value = item.TaxAmt.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(31, rowIndex)).Value = item.StatusName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(32, rowIndex)).Value = item.ConfirmUserName.ToString();

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
                        worksheet.SheetView.FreezeRows(2);
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

                    #region //刪除製令條碼QR圖片
                    DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/MoId/QR"));

                    FileInfo[] files = di.GetFiles();
                    foreach (FileInfo file in files)
                    {
                        file.Delete();
                    }
                    #endregion

                    //#region //刪除製令條碼1D圖片
                    //DirectoryInfo di2 = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/MoId/1D"));

                    //FileInfo[] files2 = di2.GetFiles();
                    //foreach (FileInfo file in files2)
                    //{
                    //    file.Delete();
                    //}
                    //#endregion

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

        #region //API
        #region //OspAlertMamo //託外逾時未加工MAMO通知
        [HttpPost]
        [Route("api/OutsourcingProduction/OspAlertMamo")]

        public void OspAlertMamo(string Company, string SecretKey/*, string ChannelId*/)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "OspMAMO");
                #endregion

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.OspAlertMamo(Company);
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

        #region //OspInAlertMamo //託外回廠前MAMO通知
        [HttpPost]
        [Route("api/OutsourcingProduction/OspInAlertMamo")]

        public void OspInAlertMamo(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "OspMAMO");
                #endregion

                #region //Request
                outsourcingProductionDA = new OutsourcingProductionDA();
                dataRequest = outsourcingProductionDA.OspInAlertMamo(Company);
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