using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using System.IO;
using System.Net;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using MESDA;
using NiceLabel.SDK;
using System.Reflection;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Threading;
using System.Threading.Tasks;

namespace Business_Manager.Controllers
{
    public class InventoryController : WebController
    {
        private InventoryDA inventoryDA = new InventoryDA();
        private ILabel label;

        #region //View
        public ActionResult MaterialRequisition()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MoWarehouseEnryManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BarcodeMerge()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult MaterialFeedingManagement()
        {
            //投料管理
            ViewLoginCheck();

            return View();
        }

        public ActionResult MaterialInputStatisticsReport()
        {
            //投料统计报表
            ViewLoginCheck();

            return View();
        }

        #endregion

        #region //Get
        #region //GetMaterialRequisition 取得領退料單資料
        [HttpPost]
        public async Task GetMaterialRequisition(int MrId = -1, string RequesitionNo = "", string MrErpFullNo = "", string DocType = "", string JournalStatus = ""
            , string NegativeStatus = "", string SignupStatus = "", string SourceType = "", string ConfirmStatus = "", string UserNo = "", string TransferStatus = "", string WoErpFullNo = ""
            , string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = await inventoryDA.GetMaterialRequisition(MrId, RequesitionNo, MrErpFullNo, DocType, JournalStatus
                    , NegativeStatus, SignupStatus, SourceType, ConfirmStatus, UserNo, TransferStatus, WoErpFullNo
                    , MtlItemNo, MtlItemName
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

        #region //GetMrDetail 取得領退料單詳細資料
        [HttpPost]
        public void GetMrDetail(int MrDetailId = -1, int MrId = -1, string WoErpFullNo = "", string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMrDetail(MrDetailId, MrId, WoErpFullNo, MtlItemNo
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

        #region //GetMrWipOrder 取得領退料單設定
        [HttpPost]
        public void GetMrWipOrder(int MrId = -1, int MoId = -1, int WoId = -1, string MrErpFullNo = "", string WoErpFullNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMrWipOrder(MrId, MoId, WoId, MrErpFullNo, WoErpFullNo
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

        #region //GetLineSeq 取得領退料單輸入序號
        [HttpPost]
        public void GetLineSeq(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetLineSeq(MrId);
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

        #region //GetMrSequence 取得領退料單領料序號
        [HttpPost]
        public void GetMrSequence(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMrSequence(MrId);
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

        #region //GetMrBarcodeRegister 取得領料條碼限制資料
        [HttpPost]
        public void GetMrBarcodeRegister(int BarcodeRegisterId = -1, int MrDetailId = -1, string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMrBarcodeRegister(BarcodeRegisterId, MrDetailId, BarcodeNo
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

        #region //GetMrBarcodeReRegister 取得領料條碼限制資料
        [HttpPost]
        public void GetMrBarcodeReRegister(int BarcodeReRegisterId = -1, int MrDetailId = -1, string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMrBarcodeReRegister(BarcodeReRegisterId, MrDetailId, BarcodeNo
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

        #region //GetErpInventoryQty 取得ERP庫存資料
        [HttpPost]
        public void GetErpInventoryQty(string MtlItemNo = "", int InventoryId = -1, string InventoryNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read,constrained-data");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetErpInventoryQty(MtlItemNo, InventoryId, InventoryNo
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

        #region //GetExcessFlag 取得領料單別是否可以超領的欄位資料
        [HttpPost]
        public void GetExcessFlag(string MrErpPrefix = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetExcessFlag(MrErpPrefix);
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

        #region //GetCheckInventoryQty 依據指定製令確認ERP庫存是否足夠
        [HttpPost]
        public void GetCheckInventoryQty(int MoId = -1, string MrType = "", double Quantity = -1, int InventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetCheckInventoryQty(MoId, MrType, Quantity, InventoryId);
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

        #region //GetBomSubstitution 取得取替代料資料
        [HttpPost]
        public void GetBomSubstitution(string BomMtlItemNo = "", string BillMtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetBomSubstitution(BomMtlItemNo, BillMtlItemNo
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

        #region //GetMoWarehouseEnry 取得入庫單單頭資料
        [HttpPost]
        public void GetMoWarehouseEnry(int MweId = -1, string MweErpPrefix = "", string MweErpFullNo = "", string DocDate = "", string ReceiptDate = ""
            , string StartDocDate = "", string EndDocDate = "", string StartReceiptDate = "", string EndReceiptDate = ""
            , string ConfirmStatus = "", string TransferStatus = ""
            , string WoErpPrefix = "", string WoErpFullNo = "", string MtlItemNo = "", string MtlItemName = "", int InventoryId = -1, string CreateBy = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "read");

                #region //Request
                dataRequest = inventoryDA.GetMoWarehouseEnry(MweId, MweErpPrefix, MweErpFullNo, DocDate, ReceiptDate
                    , StartDocDate, EndDocDate, StartReceiptDate, EndReceiptDate
                    , ConfirmStatus, TransferStatus
                    , WoErpPrefix, WoErpFullNo, MtlItemNo, MtlItemName, InventoryId, CreateBy
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

        #region //GetMweDetail 取得入庫單單身資料
        [HttpPost]
        public void GetMweDetail(int MweDetailId = -1, int MweId = -1, int MoId = -1, string MtlItemNo = "", int InventoryId = -1, string StartAcceptanceDate = "", string EndAcceptanceDate = "", string ConfirmStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                dataRequest = inventoryDA.GetMweDetail(MweDetailId, MweId, MoId, MtlItemNo, InventoryId, StartAcceptanceDate, StartAcceptanceDate, ConfirmStatus
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

        #region //GetMweDetailListData 取得入庫單單身資料(簡單)
        [HttpPost]
        public void GetMweDetailListData(int MweId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                dataRequest = inventoryDA.GetMweDetailListData(MweId);
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

        #region //GetMoIdBarcode 取得製令下的條碼
        [HttpPost]
        public void GetMoIdBarcode(int MoId = -1, int MweDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                dataRequest = inventoryDA.GetMoIdBarcode(MoId, MweDetailId);
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

        #region //GetMoIdBarcodePass 取得製令下的條碼(良品)
        [HttpPost]
        public void GetMoIdBarcodePass(int MoId = -1, int MweDetailId = -1, string BarcodeNo = "", string CurrentProdStatus = "", string BarcodeStatus = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                dataRequest = inventoryDA.GetMoIdBarcodePass(MoId, MweDetailId, BarcodeNo, CurrentProdStatus, BarcodeStatus);
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

        #region //GetMweBarcodeTransfer 取得入庫單移轉條碼
        [HttpPost]
        public void GetMweBarcodeTransfer(int MoId = -1, int MweDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                dataRequest = inventoryDA.GetMweBarcodeTransfer(MoId, MweDetailId);
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

        #region //GetWipOrderRequisitionSetQty 檢查製令是否領料+取得製令下的條碼
        [HttpPost]
        public void GetWipOrderRequisitionSetQty(int MoId = -1, int MweId = -1, int MweDetailId = -1, string ConfirmStatus = "", string MoIdType = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetWipOrderRequisitionSetQty(MoId, MweId, MweDetailId, ConfirmStatus, MoIdType);
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

        #region //GetCheckBarcode 產品入庫檢查產品條碼(刷取用)
        [HttpPost]
        public void GetCheckBarcode(int MoId = -1, string BarcodeNo = "", int InventoryId = -1, string BarcodeType = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetCheckBarcode(MoId, BarcodeNo, InventoryId, BarcodeType);
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

        #region //GetSupplierNo 取得加工廠商編號
        [HttpPost]
        public void GetSupplierNo(int SupplierId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetSupplierNo(SupplierId);
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

        #region //GetCheckMoIdType 檢查MES製令是否有條碼
        [HttpPost]
        public void GetCheckMoIdType(int MoId = -1, int MweDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetCheckMoIdType(MoId, MweDetailId);
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

        #region //GetCheckScrapRegister 檢查MES製令是否報廢品是否要入報廢倉
        [HttpPost]
        public void GetCheckScrapRegister(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetCheckScrapRegister(MoId);
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

        #region //GetMweBarcode 取得入庫單條碼相關資料
        [HttpPost]
        public void GetMweBarcode(int MweId = -1, int MweDetailId = -1, int MweBarcodeId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMweBarcode(MweId, MweDetailId, MweBarcodeId
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

        #region //GetReConfirmStatus 檢查入庫單是否有反確認過
        [HttpPost]
        public void GetReConfirmStatus(int MweId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetReConfirmStatus(MweId);
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

        #region //GetCheckoutStatus 檢查入庫單結帳碼
        [HttpPost]
        public void GetCheckoutStatus(int MweId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetCheckoutStatus(MweId);
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

        #region //GetBarcodeQty 取得條碼數量
        [HttpPost]
        public void GetBarcodeQty(string BarcodeNo = "")
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetBarcodeQty(BarcodeNo);
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

        #region //GetBarcodeStatusContent 取得條碼階段+狀態
        [HttpPost]
        public void GetBarcodeStatusContent(string Type = "", string StatusSchema = "")
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetBarcodeStatusContent(Type, StatusSchema);
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

        #region //GetPackageBarcode 取得條碼數量
        [HttpPost]
        public void GetPackageBarcode(string PackageBarcodeNo = "")
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetPackageBarcode(PackageBarcodeNo);
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

        #region //GetBarcodePrintInfo 取得拆併盤條碼資訊
        [HttpPost]
        public void GetBarcodePrintInfo(string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("BarcodeMerge", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetBarcodePrintInfo(BarcodeNo);
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

        #region //GetMergeLog 取得拆併歷史資料 --Gpai 20230511
        [HttpPost]
        public void GetMergeLog(string WoErpNo = "", string FromBarcodeNo = "", string ToBarcodeNo = "", string MtlItemNo = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("BarcodeMerge", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMergeLog(WoErpNo, FromBarcodeNo, ToBarcodeNo, MtlItemNo, StartDate, EndDate, OrderBy, PageIndex, PageSize);
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

        #region //GetLotNumberInventory 取得品號批號庫存數量
        [HttpPost]
        public void GetLotNumberInventory(string MtlItemNo = "", int InventoryId = -1, string Receipt = "", string IsNon = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetLotNumberInventory(MtlItemNo, InventoryId, Receipt, IsNon, OrderBy, PageIndex, PageSize);
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

        #region //GetCheckMoIdDateQa 檢查每日入庫型的製令品質判定結果
        [HttpPost]
        public void GetCheckMoIdDateQa(int MweId = -1, int MoId = -1,string LotNumber = "", string CheckMode = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "read");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetCheckMoIdDateQa(MweId, MoId, LotNumber, CheckMode);
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

        #region //GetPackageInBarcodes 取得包裝條碼下所有條碼
        [HttpPost]
        public void GetPackageInBarcodes(string BarcodeNo = "")
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetPackageInBarcodes(BarcodeNo);
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

        #region //GetMaterialFeedingManagement 取得投料管理下所有投料记录
        [HttpPost]
        public void GetMaterialFeedingManagement(int MatFeedRegId = -1, int MachineId = -1, int MoId = -1, string StartFeedDate = "", string EndFeedDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMaterialFeedingManagement(MatFeedRegId, MachineId, MoId, StartFeedDate, EndFeedDate
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

        #region //GetMachine 取得機台資訊
        [HttpPost]
        public void GetMachine(int MachineId = -1)
        {
            try
            {
               
                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMachine(MachineId);
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

        #region //GetMaterialInputStatisticsReport 投料统计表
        [HttpPost]
        public void GetMaterialInputStatisticsReport(string WoErpPrefix = "", string WoErpMtlItemName = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", string FilterType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMaterialInputStatisticsReport(WoErpPrefix, WoErpMtlItemName, MtlItemNo, StartDate, EndDate, FilterType
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

        #region //Add
        #region //AddMaterialRequisition 新增領退料單資料
        [HttpPost]
        public void AddMaterialRequisition(string RequesitionNo = "", string MrErpPrefix = "", string MrDate = "", string DocDate = "", string DocType = "", string Remark = "", string JournalStatus = ""
            , string SignupStatus = "", string SourceType = "", string PriorityType = "", string NegativeStatus = "", string ProductionLine = "", string ContractManufacturer = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMaterialRequisition(RequesitionNo, MrErpPrefix, MrDate, DocDate, DocType, Remark, JournalStatus
                    , SignupStatus, SourceType, PriorityType, NegativeStatus, ProductionLine, ContractManufacturer);
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

        #region //AddMrWipOrder 新增領退料單設定資料(MrWipOrder + MrDetail)
        [HttpPost]
        public void AddMrWipOrder(int MrId = -1, int MoId = -1, string WoErpPrefix = "", string WoErpNo = "", string MrType = "", double Quantity = -1
            , int InventoryId = -1, string RequisitionCode = "", string NegativeStatus = "", string Remark = ""
            , string MaterialCategory = "", string SubinventoryType = "", string LineSeq = "", int UomId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMrWipOrder(MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                    , InventoryId, RequisitionCode, NegativeStatus, Remark
                    , MaterialCategory, SubinventoryType, LineSeq, UomId);
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

        #region //AddMrDetail 新增領退料單詳細資料(MrDetail)
        [HttpPost]
        public void AddMrDetail(int MrId = -1, int MoId = -1, string MrSequence = "", int MtlItemId = -1, double Quantity = -1, int UomId = -1, int InventoryId = -1, string LotNumber = "", string DetailDesc = "", string Remark = "", string MaterialCategory = ""
            , string ProjectCode = "", string BondedStatus = "", string SubstituteStatus = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMrDetail(MrId, MoId, MrSequence, MtlItemId, Quantity, UomId, InventoryId, LotNumber, DetailDesc, Remark, MaterialCategory
                    , ProjectCode, BondedStatus, SubstituteStatus);
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

        #region //AddMrBarcodeRegister 新增物料條碼控管資料
        [HttpPost]
        public void AddMrBarcodeRegister(int MrDetailId = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMrBarcodeRegister(MrDetailId, BarcodeNo);
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

        #region //AddMrBarcodeReRegister 新增物料退料條碼控管資料
        [HttpPost]
        public void AddMrBarcodeReRegister(int MrDetailId = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMrBarcodeReRegister(MrDetailId, BarcodeNo);
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

        #region //AddKeyenceMrBarcodeRegister 新增物料條碼控管資料(Keyence模式)
        [HttpPost]
        public void AddKeyenceMrBarcodeRegister(int MrDetailId = -1, string keyenceBarcodeList = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddKeyenceMrBarcodeRegister(MrDetailId, keyenceBarcodeList);
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

        #region //AddMoWarehouseEnry 新增入庫單資料
        [HttpPost]
        public void AddMoWarehouseEnry(string DocDate = "", string ReceiptDate = "", string MweErpPrefix = "", string Remark = "", int SupplierId = -1, string ReserveTaxCode = ""
            , string TaxCode = "", string TaxType = "", double TaxRate = -1, string ApplyYYMM = "", string InvoiceType = "", string DeductType = "", string UiNo = "", string InvoiceDate = "", string InvoiceNo = ""
            , string DepartmentNo = "", string CurrencyCode = "", double Exchange = -1, int RowCnt = -1, string SupplierSo = "", string PaymentTerm = "", string AutoMaterialBilling = "", string SupplierPicking = ""
            , int ViewCompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMoWarehouseEnry(DocDate, ReceiptDate, MweErpPrefix, Remark, SupplierId, ReserveTaxCode
                    , TaxCode, TaxType, TaxRate, ApplyYYMM, InvoiceType, DeductType, UiNo, InvoiceDate, InvoiceNo
                    , DepartmentNo, CurrencyCode, Exchange, RowCnt, SupplierSo, PaymentTerm, AutoMaterialBilling, SupplierPicking
                    , ViewCompanyId);
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

        #region //AddMweDetail 新增入庫單單身資料
        [HttpPost]
        public void AddMweDetail(int MweId = -1, int MoId = -1, int MtlItemId = -1, int UomId = -1, int InventoryId = -1, string AcceptanceDate = ""
            , int ReceiptQty = -1, int AcceptQty = -1, int AvailableQty = -1, int ScriptQty = -1, int ReturnQty = -1
            , double OrigUnitPrice = -1, double OrigAmount = -1, double OrigDiscountAmt = -1, double ReceiptExpense = -1, double OrigPreTaxAmt = -1
            , double OrigTaxAmt = -1, double PreTaxAmt = -1, double TaxAmt = -1, string DiscountDescription = ""
            , string Remark = "", string ProjectCode = "", string ProcessCode = ""
            , string QcStatus = "", string Overdue = "", string ReserveTaxCode = "", string PaymentHold = ""
            , string BarcodeNoY = "", string BarcodeNoN = "", string ScrapRegister = "", string LotNumber = "", string AvailableDate = "", string ReCheckDate = ""
            , string BarcodeListAmortizeTtypeY = "", string IsAmortize = "", int AmortizeNewCount = -1, string QaResult = "", int QcRecordId = -1, string OverMsg = ""
            )
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMweDetail(MweId, MoId, MtlItemId, UomId, InventoryId, AcceptanceDate
                    , ReceiptQty, AcceptQty, AvailableQty, ScriptQty, ReturnQty
                    , OrigUnitPrice, OrigAmount, OrigDiscountAmt, ReceiptExpense, OrigPreTaxAmt
                    , OrigTaxAmt, PreTaxAmt, TaxAmt, DiscountDescription
                    , Remark, ProjectCode, ProcessCode
                    , QcStatus, Overdue, ReserveTaxCode, PaymentHold
                    , BarcodeNoY, BarcodeNoN, ScrapRegister, LotNumber, AvailableDate, ReCheckDate
                    , BarcodeListAmortizeTtypeY, IsAmortize, AmortizeNewCount, QaResult, QcRecordId, OverMsg
                    );
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

        #region //AddMweToERP 新增入庫單拋轉ERP
        [HttpPost]
        public void AddMweToERP(int MweId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMweToERP(MweId);
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

        #region //AddMweBarcodePrint 新增入庫單批量條碼
        [HttpPost]
        public void AddMweBarcodePrint(int MweDetailId = -1, int BarcodeQty = -1, string BarcodePrefix = "", string BarcodePostfix = "", int SequenceLen = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMweBarcodePrint(MweDetailId, BarcodeQty, BarcodePrefix, BarcodePostfix, SequenceLen);
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

        #region //AddEmptyMergeBarcode 新增拆併盤空條碼
        [HttpPost]
        public void AddEmptyMergeBarcode(int MoId = -1, string MtlItemNo = "", string ItemNo = "", string ItemValue = "", int CurrentMoProcessId = -1, int MoSettingId = -1, int SequenceLen = -1)
        {
            try
            {
                WebApiLoginCheck("BarcodeMerge", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddEmptyMergeBarcode(MoId, MtlItemNo, ItemNo, ItemValue, CurrentMoProcessId, MoSettingId, SequenceLen);
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

        #region //AddBatchMrWipOrder 批量新增領退料單設定資料(MrWipOrder + MrDetail)
        [HttpPost]
        public void AddBatchMrWipOrder(int MrId = -1, string MoIds = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddBatchMrWipOrder(MrId, MoIds);
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

        #region //AddMaterialFeedingManagement 新增投料记录
        [HttpPost]
        public void AddMaterialFeedingManagement(string FeedDate ="", int MachineId = -1, int MoId = -1, int MaterialId = -1, double MatInRegQty = -1, int UomId = -1, string Remarks="")
        {
            try
            {
                WebApiLoginCheck("MaterialFeedingManagement", "add");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.AddMaterialFeedingManagement(FeedDate,MachineId, MoId, MaterialId, MatInRegQty, UomId, Remarks);
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
        #region //UpdateMaterialRequisition 更新領退料單資料
        [HttpPost]
        public void UpdateMaterialRequisition(int MrId = -1, string RequesitionNo = "", string MrErpPrefix = "", string MrDate = "", string DocDate = "", string DocType = "", string Remark = "", string JournalStatus = ""
            , string SignupStatus = "", string SourceType = "", string PriorityType = "", string NegativeStatus = "", string ProductionLine = "", string ContractManufacturer = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMaterialRequisition(MrId, RequesitionNo, MrErpPrefix, MrDate, DocDate, DocType, Remark, JournalStatus
                    , SignupStatus, SourceType, PriorityType, NegativeStatus, ProductionLine, ContractManufacturer);
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

        #region //UpdateMrWipOrder 更新領退料單設定資料(重計單身)
        [HttpPost]
        public void UpdateMrWipOrder(int MrId = -1, int MoId = -1, string WoErpPrefix = "", string WoErpNo = "", string MrType = "", double Quantity = -1
            , int InventoryId = -1, string RequisitionCode = "", string NegativeStatus = "", string Remark = ""
            , string MaterialCategory = "", string SubinventoryType = "", string LineSeq = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrWipOrder(MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                    , InventoryId, RequisitionCode, NegativeStatus, Remark
                    , MaterialCategory, SubinventoryType, LineSeq);
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

        #region //UpdateMrWipOrderNoRC 更新領退料單設定資料(不重計單身)
        [HttpPost]
        public void UpdateMrWipOrderNoRC(int MrId = -1, int MoId = -1, string WoErpPrefix = "", string WoErpNo = "", string MrType = "", double Quantity = -1
            , int InventoryId = -1, string RequisitionCode = "", string NegativeStatus = "", string Remark = ""
            , string MaterialCategory = "", string SubinventoryType = "", string LineSeq = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrWipOrderNoRC(MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                    , InventoryId, RequisitionCode, NegativeStatus, Remark
                    , MaterialCategory, SubinventoryType, LineSeq);
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

        #region //UpdateMrDetail 更新領退料單詳細資料
        [HttpPost]
        public void UpdateMrDetail(int MrDetailId = -1, int MoId = -1, int WoDetailId = -1, string MrSequence = "", int MtlItemId = -1, double Quantity = -1, int UomId = -1, int InventoryId = -1, string LotNumber = "", string DetailDesc = "", string Remark = "", string MaterialCategory = ""
            , string ProjectCode = "", string BondedStatus = "", string SubstituteStatus = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrDetail(MrDetailId, MoId, MrSequence, MtlItemId, Quantity, UomId, InventoryId, LotNumber, DetailDesc, Remark, MaterialCategory
                    , ProjectCode, BondedStatus, SubstituteStatus);
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

        #region //UpdateMrLotNumber 更新領退料批號
        [HttpPost]
        public void UpdateMrLotNumber(int MrDetailId = -1, string LotNumber = "", string InventoryNo = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrLotNumber(MrDetailId, LotNumber, InventoryNo);
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

        #region //UpdateMrBomSubstitution 使用取替代料取代原領料詳細資料
        [HttpPost]
        public void UpdateMrBomSubstitution(int MrDetailId = -1, int SubMtlItemId = -1, double Quantity = -1, int BomMtlItemId = -1, double SubQuantity = -1, int InventoryId = -1, int UomId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrBomSubstitution(MrDetailId, SubMtlItemId, Quantity, BomMtlItemId, SubQuantity, InventoryId, UomId);
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

        #region //UpdateMrTransfer 將MES領料單上拋ERP，但不核單
        [HttpPost]
        public void UpdateMrTransfer(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrTransfer(MrId);
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

        #region //ConfirmMaterialRequisition 核准領退料單據
        [HttpPost]
        public void ConfirmMaterialRequisition(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "confirm");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.ConfirmMaterialRequisition(MrId);
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

        #region //UpdateMrReConfirm 領退料單反確認
        [HttpPost]
        public void UpdateMrReConfirm(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "reconfirm");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrReConfirm(MrId);
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

        #region //UpdateMrVoid 領退料單作廢
        [HttpPost]
        public void UpdateMrVoid(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "void");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrVoid(MrId);
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

        #region //UpdateMrBarcodeRegister 更新物料條碼控管資料
        [HttpPost]
        public void UpdateMrBarcodeRegister(int BarcodeRegisterId = -1, int MrDetailId = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrBarcodeRegister(BarcodeRegisterId, MrDetailId, BarcodeNo);
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

        #region //ConfirmInventoryTransaction 拋單+核准庫存異動單據
        [HttpPost]
        public void ConfirmInventoryTransaction(int ItId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.ConfirmInventoryTransaction(ItId);
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

        #region //ReConfirmInventoryTransaction 反確認庫存異動單據
        [HttpPost]
        public void ReConfirmInventoryTransaction(int ItId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.ReConfirmInventoryTransaction(ItId);
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

        #region //ConfirmItMaterialRequisition 拋單+核准領退料單據(庫存異動平台)
        [HttpPost]
        public void ConfirmItMaterialRequisition(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //拋轉
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrTransfer(MrId);

                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //核單
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.ConfirmMaterialRequisition(MrId);

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

        #region //ReConfirmItMaterialRequisition 反確認領退料單據(庫存異動平台)
        [HttpPost]
        public void ReConfirmItMaterialRequisition(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMrReConfirm(MrId);
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

        #region //UpdateItVoid 作廢庫存異動單據
        [HttpPost]
        public void UpdateItVoid(int ItId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "void");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateItVoid(ItId);
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

        #region //ConfirmTempShippingNote 核准暫出單
        [HttpPost]
        public void ConfirmTempShippingNote(int TsnId = -1)
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "confirm");

                #region //核單
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.ConfirmTempShippingNote(TsnId);
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

        #region //ReConfirmTempShippingNote 反確認暫出單
        [HttpPost]
        public void ReConfirmTempShippingNote(int TsnId = -1)
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "reconfirm");

                #region //核單
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.ReConfirmTempShippingNote(TsnId);
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


        #region //UpdateMoWarehouseEnry 更新MES入庫單單頭資料
        [HttpPost]
        public void UpdateMoWarehouseEnry(int MweId = -1, string Remark = "", string DocDate = "", string InvoiceDate = "", string InvoiceNo = "", string DepartmentNo = "", string DeductType = ""
            , string SupplierSo = "", int RowCnt = -1, string PaymentTerm = "", string AutoMaterialBilling = "", string SupplierPicking = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMoWarehouseEnry(MweId, Remark, DocDate, InvoiceDate, InvoiceNo, DepartmentNo, DeductType
                    , SupplierSo, RowCnt, PaymentTerm, AutoMaterialBilling, SupplierPicking);
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

        #region //UpdateMweDetail 更新MES入庫單單身資料
        [HttpPost]
        public void UpdateMweDetail(int MweDetailId = -1, int MweId = -1, int InventoryId = -1, string AcceptanceDate = ""
            , int ReceiptQty = -1, int AcceptQty = -1, int AvailableQty = -1, int ScriptQty = -1, int ReturnQty = -1
            , double OrigUnitPrice = -1, double OrigAmount = -1, double OrigDiscountAmt = -1, double ReceiptExpense = -1, string DiscountDescription = "", double OrigPreTaxAmt = -1, double OrigTaxAmt = -1
            , double PreTaxAmt = -1, double TaxAmt = -1, string Remark = "", string ProjectCode = "", string ProcessCode = "", string QcStatus = ""
            , string Overdue = "", string ReserveTaxCode = "", string PaymentHold = "", string MoIdType = "", string BarcodeNoY = "", string BarcodeNoN = "", string ScrapRegister = "", string LotNumber = ""
            , string BarcodeListAmortizeTtypeY = "", int AmortizeNewCount = -1, string IsAmortize = "", string QaResult = "", int QcRecordId = -1, string OverMsg = ""
            )
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMweDetail(MweDetailId, MweId, InventoryId, AcceptanceDate, ReceiptQty, AcceptQty, AvailableQty, ScriptQty, ReturnQty,
                    OrigUnitPrice, OrigAmount, OrigDiscountAmt, ReceiptExpense, DiscountDescription, OrigPreTaxAmt, OrigTaxAmt
                    , PreTaxAmt, TaxAmt, Remark, ProjectCode, ProcessCode, QcStatus, Overdue, ReserveTaxCode, PaymentHold, MoIdType
                    , BarcodeNoY, BarcodeNoN, ScrapRegister, LotNumber
                    , BarcodeListAmortizeTtypeY, AmortizeNewCount, IsAmortize, QaResult, QcRecordId, OverMsg
                    );
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

        #region //UpdateBarcodeQty 更新MES入庫單單身資料
        [HttpPost]
        public void UpdateBarcodeQty(int MweDetailId, string BarcodeListOre = "", string BarcodeListAdd = "", int BarcodeAllQty = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateBarcodeQty(MweDetailId, BarcodeListOre, BarcodeListAdd, BarcodeAllQty);
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

        #region //UpdateMweERPConfirm MES入庫單單據核單(ERP核單)
        [HttpPost]
        public void UpdateMweERPConfirm(int MweId = -1, int ConfirmModel = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "confirm");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMweERPConfirm(MweId, ConfirmModel);
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

        #region //UpdateMweReconfirm 入庫單反確認
        [HttpPost]
        public void UpdateMweReconfirm(int MweId = -1, int ConfirmModel = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "reconfirm");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMweReconfirm(MweId, ConfirmModel);
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

        #region //UpdateReConfirmStatus 入庫單重新編輯
        [HttpPost]
        public void UpdateReConfirmStatus(int MweId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateReConfirmStatus(MweId);
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

        #region //UpdateMweVoid 入庫單作廢
        [HttpPost]
        public void UpdateMweVoid(int MweId = -1, int OperateUserId = -1, int OperateCompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "void");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.UpdateMweVoid(MweId, OperateUserId, OperateCompanyId);
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

        #region //SplitMergeBarcodePrint 拆併盤
        [HttpPost]
        public void SplitMergeBarcodePrint(string FirstBarcode = "", string SecondBarcode = "", int Qty = -1)
        {
            try
            {
                WebApiLoginCheck("BarcodeMerge", "update");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.SplitMergeBarcodePrint(FirstBarcode, SecondBarcode, Qty);
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
        #region //DeleteMaterialRequisition -- 刪除領退料單資料
        [HttpPost]
        public void DeleteMaterialRequisition(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "delete");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMaterialRequisition(MrId);
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

        #region //DeleteMrWipOrder -- 刪除領退料單設定資料
        [HttpPost]
        public void DeleteMrWipOrder(int MrId = -1, int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "delete");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMrWipOrder(MrId, MoId);
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

        #region //DeleteMrDetail -- 刪除領退料單詳細資料
        [HttpPost]
        public void DeleteMrDetail(int MrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "delete");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMrDetail(MrDetailId);
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

        #region //BatchDeleteMrDetail -- 批量刪除領退料單詳細資料
        [HttpPost]
        public void BatchDeleteMrDetail(string MrDetailList = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "delete");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.BatchDeleteMrDetail(MrDetailList);
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

        #region //DeleteMrBarcodeRegister -- 刪除物料條碼控管資料
        [HttpPost]
        public void DeleteMrBarcodeRegister(int BarcodeRegisterId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMrBarcodeRegister(BarcodeRegisterId);
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

        #region //DeleteMrBarcodeReRegister -- 刪除物料退料條碼控管資料
        [HttpPost]
        public void DeleteMrBarcodeReRegister(int BarcodeReRegisterId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMrBarcodeReRegister(BarcodeReRegisterId);
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

        #region //DeleteMoWarehouseEnry -- 刪除入庫單
        [HttpPost]
        public void DeleteMoWarehouseEnry(int MweId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "delete");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMoWarehouseEnry(MweId);
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

        #region //DeleteMweDetail -- 刪除入庫單單身
        [HttpPost]
        public void DeleteMweDetail(int MweId = -1, int MweDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMweDetail(MweId, MweDetailId);
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

        #region //DeleteBarcodeAll -- 刪除入庫單單身綁定的條碼(全部)
        [HttpPost]
        public void DeleteBarcodeAll(int MweId = -1, int MweDetailId = -1, string BarcodeId = "")
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteBarcodeAll(MweId, MweDetailId, BarcodeId);
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

        #region //DeleteMweBarcodePrint -- 刪除數量入庫單身的批量條碼
        [HttpPost]
        public void DeleteMweBarcodePrint(int MweBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMweBarcodePrint(MweBarcodeId);
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

        #region //DeleteMweBarcodePrintAll -- 刪除數量入庫單身的批量條碼(全部)
        [HttpPost]
        public void DeleteMweBarcodePrintAll(int MweDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MoWarehouseEnryManagment", "detail");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.DeleteMweBarcodePrintAll(MweDetailId);
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

        #region //EXCEL API
        [HttpPost]
        public FileContentResult ExportExcel()
        {
            System.IO.Stream spreadSheetStream = new System.IO.MemoryStream();
            XLWorkbook workbook = new XLWorkbook(); //載入EXCEL
            IXLWorksheet worksheet = workbook.Worksheets.Add("example"); //新增工作表
            //IXLWorksheet worksheet = workbook.Worksheet("工作表1"); //使用【工作表1】作為worksheet

            worksheet.Cell(1, 1).SetValue("TEST TABLE"); //設定第一列第一行
            worksheet.Cell(2, 1).SetValue("TEST TABLE2");//設定第二列第一行

            string range2 = "A1:H2";
            worksheet.Range(range2).Merge(); //設定合併儲存格 A1~H2
            worksheet.Cell(1, 1).Value = "Employee Report"; //設定第一列第一行資料

            //worksheet.Column(1).Width = 50; //設定第一欄寬度為50
            //worksheet.Column(1).AdjustToContents(); //設定第一欄寬度自動變化
            worksheet.Columns().AdjustToContents(); //設定所有欄位寬度自適應

            worksheet.Range("A1").Style.Font.SetFontSize(18); //設定A1字體大小

            worksheet.Cell(1, 1).Style.Fill.SetBackgroundColor(XLColor.AppleGreen); //設定第一列第一行欄位背景顏色

            worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //水平置中
            worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; //垂直置中

            //測試迴圈動態新增
            int i = 0;
            for (i = 3; i <= 50; i++)
            {
                worksheet.Cell(i, 1).SetValue("第" + i.ToString() + "列第1行");
            }
            //測試迴圈動態新增 END

            //設定範圍內邊框樣式
            string range = "A1:H" + i.ToString();
            worksheet.Range(range).Style.Font.FontName = "微軟正黑體"; //設定字型
            worksheet.Range("A1:H2").Style.Font.SetFontColor(XLColor.White); //設定範圍內字體顏色
            worksheet.Range(range).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Range(range).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range(range).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.Range(range).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Range(range).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            //設定範圍內邊框樣式 END

            workbook.SaveAs(spreadSheetStream);
            spreadSheetStream.Position = 0;

            var memoryStream = new MemoryStream();
            spreadSheetStream.CopyTo(memoryStream);

            return File(memoryStream.ToArray(), "application/vnd.ms-excel", DateTime.Now.ToString("yyy-MM-dd") + ".xlsx");
        }

        [HttpPost]
        public FileContentResult LoadExcel(string DataPath = "")
        {
            //string filePath = @"C:\Users\ZY01001\Desktop\複本 企業管理平台功能說明.xlsx";
            DataPath = @"C:\Users\ZY01001\Downloads\使用者管理Excel檔2.xlsx";
            string filePath = DataPath;
            System.IO.Stream spreadSheetStream = new System.IO.MemoryStream();
            XLWorkbook workbook = new XLWorkbook(filePath);
            //IXLWorksheet worksheet = workbook.Worksheets.Add("example");
            var ws1 = workbook.Worksheet(1);
            var data = ws1.Cell("A1").GetValue<string>();

            workbook.SaveAs(spreadSheetStream);
            spreadSheetStream.Position = 0;

            var memoryStream = new MemoryStream();
            spreadSheetStream.CopyTo(memoryStream);

            return File(memoryStream.ToArray(), "application/vnd.ms-excel", DateTime.Now.ToString("yyy-MM-dd") + ".xls");
        }
        #endregion

        #region //PDF API
        public class UnicodeFontFactory : FontFactoryImp
        {

            private static readonly string arialFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                "arialuni.ttf");//arial unicode MS是完整的unicode字型。
            private static readonly string 標楷體Path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
              "KAIU.TTF");//標楷體


            public override iTextSharp.text.Font GetFont(string fontname, string encoding, bool embedded, float size, int style, BaseColor color,
                bool cached)
            {
                //可用Arial或標楷體，自己選一個
                BaseFont baseFont = BaseFont.CreateFont(標楷體Path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                return new iTextSharp.text.Font(baseFont, size, style, color);
            }


        }

        [HttpPost]
        public FileContentResult ExportPDF()
        {
            WebClient wc = new WebClient();
            //從網址下載Html字串
            string htmlText = wc.DownloadString(@"C:\Users\ZY01001\Desktop\order_html\test.html");
            //StreamReader sr = new StreamReader(@"C:\Users\ZY01001\Desktop\order_html\order.html",
            //System.Text.Encoding.UTF8);
            byte[] pdfFile = ConvertHtmlTextToPDF(htmlText);

            return File(pdfFile, "application/pdf", "範例PDF檔.pdf");
        }

        public byte[] ConvertHtmlTextToPDF(string htmlText)
        {
            if (string.IsNullOrEmpty(htmlText))
            {
                return null;
            }
            //避免當htmlText無任何html tag標籤的純文字時，轉PDF時會掛掉，所以一律加上<p>標籤
            htmlText = "<p>" + htmlText + "</p>";

            MemoryStream outputStream = new MemoryStream();//要把PDF寫到哪個串流
            byte[] data = Encoding.UTF8.GetBytes(htmlText);//字串轉成byte[]
            MemoryStream msInput = new MemoryStream(data);
            iTextSharp.text.Document doc = new iTextSharp.text.Document();//要寫PDF的文件，建構子沒填的話預設直式A4
            PdfWriter writer = PdfWriter.GetInstance(doc, outputStream);
            //指定文件預設開檔時的縮放為100%
            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            //開啟Document文件 
            doc.Open();
            //使用XMLWorkerHelper把Html parse到PDF檔裡
            XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, msInput, null, Encoding.UTF8, new UnicodeFontFactory());
            //將pdfDest設定的資料寫到PDF檔
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, writer);
            writer.SetOpenAction(action);
            doc.Close();
            msInput.Close();
            outputStream.Close();
            //回傳PDF檔案 
            return outputStream.ToArray();

        }
        #endregion

        #region//Print

        #region //PrintMweBarcode -- 入庫單標籤列印  -- Shintokuro 2023-03-10
        [HttpPost]
        public void PrintMweBarcode(int MweId, int MweDetailId, int MweBarcodeId, string PrintMachine)
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
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintMweBarcodePrint\\Template\\LabelName.txt", Encoding.Default);


                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintMweBarcodePrint\\Template\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();

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

                #region //標籤列印


                dataRequest = inventoryDA.GetMweBarcode(MweId, MweDetailId, MweBarcodeId
                    , "", -1, -1);

                #region //Response
                string MweBarcodeNo = "";
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    MweBarcodeNo = result[0]["BarcodeNo"].ToString();

                    label.Variables["BARCODE_NO"].SetValue(MweBarcodeNo);
                    Thread.Sleep(2000);
                    label.Print(1);

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                }
                #endregion


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

        #region //PrintMweBarcodeAll -- 入庫單標籤列印(全部)  -- Shintokuro 2023-03-10
        [HttpPost]
        public void PrintMweBarcodeAll(int MweId = -1, int MweDetailId = -1, int MweBarcodeId = -1, string PrintMachine = "")
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
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintMweBarcodePrint\\Template\\LabelName.txt", Encoding.Default);


                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintMweBarcodePrint\\Template\\PrinterName.txt", Encoding.Default);
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

                #region //標籤列印


                dataRequest = inventoryDA.GetMweBarcode(MweId, MweDetailId, MweBarcodeId
                    , "", -1, -1);

                #region //Response
                string MweBarcodeNo = "";
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    for (var i = 0; i < result.Count(); i++)
                    {
                        MweBarcodeNo = result[i]["BarcodeNo"].ToString();

                        label.Variables["BARCODE_NO"].SetValue(MweBarcodeNo);
                        Thread.Sleep(2000);
                        label.Print(1);
                    }
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                }
                #endregion


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

        #region //Download
        #region //Word
        #region //MaterialRequisitionDocDownload 領料單資料單據下載
        public void MaterialRequisitionDocDownload(int MrId = -1)
        {
            try
            {
                inventoryDA.AddUserOperateLog(MrId);
                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMrDocInfo(MrId);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = "", depositFilePath = "", user = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;
                    DateTime currendate = default(DateTime);
                    string formattedDate = currendate.ToString("yyyy/MM/dd");

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("單據資料不完整!");

                    #region //判斷公司別文件
                    int CurrentCompany = Convert.ToInt32(Session["UserCompany"]);

                    if (Session["CompanySwitch"] != null)
                    {
                        if (Session["CompanySwitch"].ToString() == "manual")
                        {
                            CurrentCompany = Convert.ToInt32(Session["CompanyId"]);
                        }
                    }

                    var moctc = result["data"];
                    var mocte = result["dataDetail"];

                    switch (CurrentCompany)
                    {
                        case 2: //中揚
                            if (Convert.ToInt32(moctc[0]["MQ010"]) < 0)
                            {
                                wordFileName = "領料單-{0}-{1}";
                                filePath = "~/WordTemplate/MES/P-W001表02-01 領料單.docx";
                                secondFilePath = "~/WordTemplate/MES/P-W001表02-01 領料單 換頁.docx";
                            }
                            else if (Convert.ToInt32(moctc[0]["MQ010"]) > 0)
                            {
                                wordFileName = "退料單-{0}-{1}";
                                filePath = "~/WordTemplate/MES/P-W001表03-01 退料單.docx";
                                secondFilePath = "~/WordTemplate/MES/P-W001表03-01 退料單 換頁.docx";
                            }
                            else
                            {
                                throw new SystemException("庫存成本影響碼錯誤!!");
                            }
                            break;

                        case 3: //中揚
                            if (Convert.ToInt32(moctc[0]["MQ010"]) < 0)
                            {
                                wordFileName = "領料單-{0}-{1}";
                                filePath = "~/WordTemplate/MES/M-01-表04 領料單-紘立.docx";
                                secondFilePath = "~/WordTemplate/MES/M-01-表04 領料單-紘立 換頁.docx";
                            }
                            else if (Convert.ToInt32(moctc[0]["MQ010"]) > 0)
                            {
                                wordFileName = "退料單-{0}-{1}";
                                filePath = "~/WordTemplate/MES/M-01-表07 退料單-紘立.docx";
                                secondFilePath = "~/WordTemplate/MES/M-01-表07 退料單-紘立 換頁.docx";
                            }
                            else
                            {
                                throw new SystemException("庫存成本影響碼錯誤!!");
                            }
                            break;

                        case 4: //晶彩
                            if (Convert.ToInt32(moctc[0]["MQ010"]) < 0)
                            {
                                wordFileName = "領料單-{0}-{1}";
                                filePath = "~/WordTemplate/MES/R-PU00-05 領料單-晶彩.docx";
                                secondFilePath = "~/WordTemplate/MES/R-PU00-05 領料單-晶彩 換頁.docx";
                            }
                            else if (Convert.ToInt32(moctc[0]["MQ010"]) > 0)
                            {
                                wordFileName = "退料單-{0}-{1}";
                                filePath = "~/WordTemplate/MES/R-PU00-16 退料單-晶彩.docx";
                                secondFilePath = "~/WordTemplate/MES/R-PU00-16 退料單-晶彩 換頁.docx";
                            }
                            else
                            {
                                throw new SystemException("庫存成本影響碼錯誤!!");
                            }
                            break;
                        default:
                            throw new SystemException("目前公司別未開放功能!");
                    }
                    #endregion

                    #region //產生Doc
                    totalPage = mocte.Count() / 5 + (mocte.Count() % 5 > 0 ? 1 : 0);

                    if (totalPage == 1)
                    {
                        #region //單頁
                        using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                        {
                            wordFileName = string.Format(wordFileName, moctc[0]["TC001"].ToString(), moctc[0]["TC002"].ToString());

                            #region //單頭
                            doc.ReplaceText("[MQ002]", moctc[0]["MQ002"].ToString());
                            doc.ReplaceText("[TC001]", moctc[0]["TC001"].ToString());
                            doc.ReplaceText("[TC002]", moctc[0]["TC002"].ToString());
                            doc.ReplaceText("[TC014]", moctc[0]["TC014"].ToString());
                            doc.ReplaceText("[TC004]", moctc[0]["TC004"].ToString());
                            doc.ReplaceText("[TC005]", moctc[0]["TC005"].ToString());
                            doc.ReplaceText("[TC007]", moctc[0]["TC007"].ToString());
                            doc.ReplaceText("[MB002]", moctc[0]["MB002"].ToString());
                            doc.ReplaceText("[MD002]", moctc[0]["MD002"].ToString());
                            doc.ReplaceText("[TC006]", moctc[0]["TC006"].ToString());

                            doc.ReplaceText("[cp]", "1");
                            doc.ReplaceText("[tp]", "1");
                            doc.ReplaceText("[currendate]", formattedDate);

                            #endregion

                            #region //單身
                            if (mocte.Count() > 0)
                            {
                                for (int i = 0; i < mocte.Count(); i++)
                                {
                                    string Type = mocte[i]["Type"].ToString() == "2" ? "正式" : "工程";
                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocte[i]["Seq"].ToString());
                                    doc.ReplaceText(string.Format("[Type{0:00}]", i + 1), Type);

                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocte[i]["MtlItemNo"].ToString());

                                    string tempMtlItemName = mocte[i]["MtlItemName"].ToString();
                                    if (tempMtlItemName.Length > 33)
                                    {
                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocte[i]["MtlItemSpec"].ToString());

                                    doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocte[i]["UomNo"].ToString());

                                    doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocte[i]["Quantity"].ToString());
                                    doc.ReplaceText(string.Format("[InventoryNo{0:00}]", i + 1), mocte[i]["InventoryNo"].ToString());
                                    doc.ReplaceText(string.Format("[InventoryName{0:00}]", i + 1), mocte[i]["InventoryName"].ToString());

                                    doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocte[i]["WoErpFullNo"].ToString());
                                    doc.ReplaceText(string.Format("[ProcessCode{0:00}]", i + 1), mocte[i]["ProcessCode"].ToString());
                                    doc.ReplaceText(string.Format("[Location{0:00}]", i + 1), mocte[i]["Location"].ToString());

                                    doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());

                                    string tempRequisition = mocte[i]["Requisition"].ToString();
                                    if (tempRequisition.Length > 50)
                                    {
                                        tempRequisition = BaseHelper.StrLeft(tempRequisition, 50) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[Requisition{0:00}]", i + 1), tempRequisition);

                                    string tempRemark = mocte[i]["Remark"].ToString();
                                    if (tempRemark.Length > 50)
                                    {
                                        tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                }

                                #region //剩餘欄位
                                if (mocte.Count() < 5)
                                {
                                    for (int i = mocte.Count(); i < 5; i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Type{0:00}]", i + 1), "");

                                        if (i == mocte.Count() && i < 5)
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                        }
                                        else
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                        }

                                        doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InventoryNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InventoryName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ProcessCode{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Location{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Requisition{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");
                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("【單身資料】錯誤!");
                            }
                            #endregion

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region //多頁
                        using (DocX doc = DocX.Load(Server.MapPath(secondFilePath)))
                        {
                            wordFileName = string.Format(wordFileName, moctc[0]["TC001"].ToString(), moctc[0]["TC002"].ToString());

                            for (int p = 1; p <= totalPage; p++)
                            {
                                if (p != totalPage)
                                {
                                    #region //預設頁面
                                    if (p != 1)
                                    {
                                        using (DocX subDoc = DocX.Load(Server.MapPath(secondFilePath)))
                                        {
                                            doc.InsertDocument(subDoc, true);
                                        }
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", moctc[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TC001]", moctc[0]["TC001"].ToString());
                                    doc.ReplaceText("[TC002]", moctc[0]["TC002"].ToString());
                                    doc.ReplaceText("[TC014]", moctc[0]["TC014"].ToString());
                                    doc.ReplaceText("[TC004]", moctc[0]["TC004"].ToString());
                                    doc.ReplaceText("[TC005]", moctc[0]["TC005"].ToString());
                                    doc.ReplaceText("[TC007]", moctc[0]["TC007"].ToString());
                                    doc.ReplaceText("[MB002]", moctc[0]["MB002"].ToString());
                                    doc.ReplaceText("[MD002]", moctc[0]["MD002"].ToString());
                                    doc.ReplaceText("[TC006]", moctc[0]["TC006"].ToString());

                                    doc.ReplaceText("[cp]", p.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[NextPage]", "接下頁...");
                                    doc.ReplaceText("[currendate]", formattedDate);

                                    #endregion

                                    #region //單身
                                    if (mocte.Count() > 0)
                                    {
                                        for (int i = 0; i < 5; i++)
                                        {
                                            string Type = mocte[i + (p - 1) * 5]["Type"].ToString() == "2" ? "正式" : "工程";
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocte[i + (p - 1) * 5]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[Type{0:00}]", i + 1), Type);
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["MtlItemNo"].ToString());

                                            string tempMtlItemName = mocte[i + (p - 1) * 5]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocte[i + (p - 1) * 5]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["UomNo"].ToString());

                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocte[i + (p - 1) * 5]["Quantity"].ToString());
                                            doc.ReplaceText(string.Format("[InventoryNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["InventoryNo"].ToString());
                                            doc.ReplaceText(string.Format("[InventoryName{0:00}]", i + 1), mocte[i + (p - 1) * 5]["InventoryName"].ToString());

                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["WoErpFullNo"].ToString());
                                            doc.ReplaceText(string.Format("[ProcessCode{0:00}]", i + 1), mocte[i + (p - 1) * 5]["ProcessCode"].ToString());
                                            doc.ReplaceText(string.Format("[Location{0:00}]", i + 1), mocte[i]["Location"].ToString());

                                            doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i + (p - 1) * 5]["LotNumber"].ToString());

                                            string tempRequisition = mocte[i + (p - 1) * 5]["Requisition"].ToString();
                                            if (tempRequisition.Length > 50)
                                            {
                                                tempRequisition = BaseHelper.StrLeft(tempRequisition, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Requisition{0:00}]", i + 1), tempRequisition);

                                            string tempRemark = mocte[i + (p - 1) * 5]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion
                                    #endregion
                                }
                                else
                                {
                                    #region //最後一頁
                                    using (DocX finalDoc = DocX.Load(Server.MapPath(filePath)))
                                    {
                                        doc.InsertDocument(finalDoc, true);
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", moctc[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TC001]", moctc[0]["TC001"].ToString());
                                    doc.ReplaceText("[TC002]", moctc[0]["TC002"].ToString());
                                    doc.ReplaceText("[TC014]", moctc[0]["TC014"].ToString());
                                    doc.ReplaceText("[TC004]", moctc[0]["TC004"].ToString());
                                    doc.ReplaceText("[TC005]", moctc[0]["TC005"].ToString());
                                    doc.ReplaceText("[TC007]", moctc[0]["TC007"].ToString());
                                    doc.ReplaceText("[MB002]", moctc[0]["MB002"].ToString());
                                    doc.ReplaceText("[MD002]", moctc[0]["MD002"].ToString());
                                    doc.ReplaceText("[TC006]", moctc[0]["TC006"].ToString());

                                    doc.ReplaceText("[cp]", totalPage.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[currendate]", formattedDate);

                                    #endregion

                                    #region //單身
                                    if (mocte.Count() > 0)
                                    {
                                        for (int i = 0; i < (mocte.Count() % 5 != 0 ? mocte.Count() % 5 : 5); i++)
                                        {
                                            string Type = mocte[i + (p - 1) * 5]["Type"].ToString() == "2" ? "正式" : "工程";
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocte[i + (p - 1) * 5]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[Type{0:00}]", i + 1), Type);
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["MtlItemNo"].ToString());

                                            string tempMtlItemName = mocte[i + (p - 1) * 5]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocte[i + (p - 1) * 5]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["UomNo"].ToString());

                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocte[i + (p - 1) * 5]["Quantity"].ToString());
                                            doc.ReplaceText(string.Format("[InventoryNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["InventoryNo"].ToString());
                                            doc.ReplaceText(string.Format("[InventoryName{0:00}]", i + 1), mocte[i + (p - 1) * 5]["InventoryName"].ToString());

                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocte[i + (p - 1) * 5]["WoErpFullNo"].ToString());
                                            doc.ReplaceText(string.Format("[ProcessCode{0:00}]", i + 1), mocte[i + (p - 1) * 5]["ProcessCode"].ToString());
                                            doc.ReplaceText(string.Format("[Location{0:00}]", i + 1), mocte[i + (p - 1) * 5]["Location"].ToString());

                                            doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i + (p - 1) * 5]["LotNumber"].ToString());

                                            string tempRequisition = mocte[i + (p - 1) * 5]["Requisition"].ToString();
                                            if (tempRequisition.Length > 50)
                                            {
                                                tempRequisition = BaseHelper.StrLeft(tempRequisition, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Requisition{0:00}]", i + 1), tempRequisition);

                                            string tempRemark = mocte[i + (p - 1) * 5]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion

                                    #region //剩餘欄位
                                    if (mocte.Count() % 5 != 0)
                                    {
                                        for (int i = mocte.Count() % 5; i < 5; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Type{0:00}]", i + 1), "");

                                            if (i == mocte.Count() % 5 && i < 5)
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InventoryNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InventoryName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ProcessCode{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Location{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Requisition{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");
                                        }
                                    }
                                    #endregion
                                    #endregion
                                }
                            }

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }

                    fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = wordFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = wordFileName,
                        fileExtension = ".docx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "目前尚無資料"
                    });
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

        #region //MoWarehouseEnryDocDownload 入庫單資料單據下載
        public void MoWarehouseEnryDocDownload(int MweId = -1)
        {
            try
            {
                inventoryDA.AddUserOperateLog(MweId);

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMweDocInfo(MweId);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = "", depositFilePath = "", user = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("單據資料不完整!");

                    #region //判斷公司別文件
                    int CurrentCompany = Convert.ToInt32(Session["UserCompany"]);

                    if (Session["CompanySwitch"] != null)
                    {
                        if (Session["CompanySwitch"].ToString() == "manual")
                        {
                            CurrentCompany = Convert.ToInt32(Session["CompanyId"]);
                        }
                    }

                    var mocth = result["data"];
                    var mocti = result["dataDetail"];
                    DateTime currendate = DateTime.Now;
                    string formattedDate = currendate.ToString("yyyy/MM/dd");


                    switch (CurrentCompany)
                    {
                        case 2: //中揚
                            wordFileName = "入庫單-{0}-{1}";
                            filePath = "~/WordTemplate/MES/P-W001表04-01 入庫單.docx";
                            secondFilePath = "~/WordTemplate/MES/P-W001表04-01 入庫單 換頁.docx";
                            break;
                        case 3: //紘立
                            wordFileName = "入庫單-{0}-{1}";
                            filePath = "~/WordTemplate/MES/M-01-表06 入庫單-紘立.docx";
                            secondFilePath = "~/WordTemplate/M-01-表06 入庫單-紘立 換頁.docx";
                            break;
                        case 4: //晶彩
                            wordFileName = "入庫單-{0}-{1}";
                            filePath = "~/WordTemplate/MES/R-PU00-06 入庫單-晶彩.docx";
                            secondFilePath = "~/WordTemplate/MES/R-PU00-06 入庫單-晶彩 換頁.docx";
                            break;
                        default:
                            throw new SystemException("目前公司別未開放功能!");
                    }
                    #endregion

                    #region //產生Doc
                    totalPage = mocti.Count() / 8 + (mocti.Count() % 8 > 0 ? 1 : 0);

                    if (totalPage == 1)
                    {
                        #region //單頁
                        using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                        {
                            wordFileName = string.Format(wordFileName, mocth[0]["TH001"].ToString(), mocth[0]["TH002"].ToString());

                            #region //單頭
                            doc.ReplaceText("[MQ002]", mocth[0]["MQ002"].ToString());//=-
                            doc.ReplaceText("[TH001]", mocth[0]["TH001"].ToString());//=-
                            doc.ReplaceText("[TH002]", mocth[0]["TH002"].ToString());//=-
                            doc.ReplaceText("[TH005]", mocth[0]["TH005"].ToString());//=-
                            doc.ReplaceText("[TH006]", mocth[0]["TH006"].ToString());//=-
                            doc.ReplaceText("[TH029]", mocth[0]["TH029"].ToString());//=-
                            doc.ReplaceText("[TH003]", mocth[0]["TH003"].ToString());//=
                            doc.ReplaceText("[TH004]", mocth[0]["TH004"].ToString());//=-
                            doc.ReplaceText("[TH009]", mocth[0]["TH009"].ToString());//=-
                            doc.ReplaceText("[TH010]", mocth[0]["TH010"].ToString());//=-

                            if (CurrentCompany == 4) {
                                doc.ReplaceText("[CurrencyTotla]", mocth[0]["CurrencyTotal"].ToString());

                                doc.ReplaceText("[TH007]", mocth[0]["TH007"].ToString());//-
                                doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());//-
                                doc.ReplaceText("[TH011]", mocth[0]["TH011"].ToString());//-
                                doc.ReplaceText("[TH012]", mocth[0]["TH012"].ToString());//-
                                doc.ReplaceText("[TH013]", mocth[0]["TH013"].ToString());//-
                                doc.ReplaceText("[TH014]", mocth[0]["TH014"].ToString());//-
                                doc.ReplaceText("[TH030]", mocth[0]["TH030"].ToString());//-
                                doc.ReplaceText("[TH015]", mocth[0]["TH015"].ToString());//-
                                doc.ReplaceText("[TH016]", mocth[0]["TH016"].ToString());//-
                                doc.ReplaceText("[TH023]", mocth[0]["TH023"].ToString());//-
                                doc.ReplaceText("[TH033]", mocth[0]["TH033"].ToString());//-

                            }
                            

                            doc.ReplaceText("[UDF01]", mocth[0]["UDF01"].ToString());//=
                            doc.ReplaceText("[TH032]", mocth[0]["TH032"].ToString());
                            doc.ReplaceText("[TH018]", mocth[0]["TH018"].ToString());
                            doc.ReplaceText("[TH019]", mocth[0]["TH019"].ToString());
                            doc.ReplaceText("[TH021]", mocth[0]["TH021"].ToString());
                            doc.ReplaceText("[TH022]", mocth[0]["TH022"].ToString());
                            doc.ReplaceText("[TH027]", mocth[0]["TH027"].ToString());
                            doc.ReplaceText("[TH020]", mocth[0]["TH020"].ToString());
                            doc.ReplaceText("[TH031]", mocth[0]["TH031"].ToString());

                            doc.ReplaceText("[CurrencyTotal]", mocth[0]["CurrencyTotal"].ToString());
                            doc.ReplaceText("[LocalTotal]", mocth[0]["LocalTotal"].ToString());



                            doc.ReplaceText("[cp]", "1");
                            doc.ReplaceText("[tp]", "1");
                            doc.ReplaceText("[currendate]", formattedDate);

                            #endregion

                            #region //單身
                            if (mocti.Count() > 0)
                            {
                                for (int i = 0; i < mocti.Count(); i++)
                                {
                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocti[i]["Seq"].ToString());
                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocti[i]["MtlItemNo"].ToString());
                                    string tempMtlItemName = mocti[i]["MtlItemName"].ToString();
                                    if (tempMtlItemName.Length > 33)
                                    {
                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);
                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocti[i]["MtlItemSpec"].ToString());

                                    doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocti[i]["Quantity"].ToString());
                                    doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocti[i]["UomNo"].ToString());
                                    doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), mocti[i]["Acute"].ToString());

                                    doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AccertQty"]).ToString("F3"));
                                    doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ScrapQty"]).ToString("F3"));
                                    doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ReturnQty"]).ToString("F3"));

                                    doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), mocti[i]["InventoryNo"].ToString());
                                    doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), mocti[i]["InventoryName"].ToString());

                                    doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), mocti[i]["UnitAmount"].ToString());
                                    doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), mocti[i]["TotalAmount"].ToString());
                                    doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), mocti[i]["DeduteAmount"].ToString());
                                    doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), mocti[i]["FeeAmount"].ToString());

                                    if (CurrentCompany == 4) {
                                        doc.ReplaceText(string.Format("[InQty{0:00}]", i + 1), mocti[i]["Quantity"].ToString());
                                        doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), mocti[i]["Verify"].ToString());
                                        doc.ReplaceText(string.Format("[DeliverQty{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AccertQty"]).ToString("F3"));//已交
                                        doc.ReplaceText(string.Format("[InspectQty{0:00}]", i + 1), mocti[i]["Inspec"].ToString());//待驗
                                        doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["OverRate"]).ToString("F3") + '%');//超收
                                        doc.ReplaceText(string.Format("[AmountQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AmountQty"]).ToString("F3"));//計價
                                        doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), mocti[i]["TI018"].ToString());//驗收日
                                        doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), mocti[i]["LotNo"].ToString());//批號
                                        doc.ReplaceText(string.Format("[Projrct{0:00}]", i + 1), mocti[i]["Project"].ToString());//專案
                                        doc.ReplaceText(string.Format("[Pro{0:00}]", i + 1), mocti[i]["TI015"].ToString());//專案


                                    }

                                    doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocti[i]["WoErpFullNo"].ToString());
                                    //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());
                                    string tempRemark = mocti[i]["Remark"].ToString();
                                    if (tempRemark.Length > 50)
                                    {
                                        tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                }

                                #region //剩餘欄位
                                if (mocti.Count() < 8)
                                {
                                    for (int i = mocti.Count(); i < 8; i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                        if (i == mocti.Count() && i < 8)
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                        }
                                        else
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                        }
                                        doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), "");
                                        //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());
                                        doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");

                                        if (CurrentCompany == 4)
                                        {
                                            doc.ReplaceText(string.Format("[InQty{0:00}]", i + 1), "");

                                            doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[DeliverQty{0:00}]", i + 1), "");//已交
                                            doc.ReplaceText(string.Format("[InspectQty{0:00}]", i + 1), "");//待驗
                                            doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), "");//超收
                                            doc.ReplaceText(string.Format("[AmountQry{0:00}]", i + 1), "");//計價
                                            doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), "");//驗收日
                                            doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), "");//批號
                                            doc.ReplaceText(string.Format("[Projrct{0:00}]", i + 1), "");//專案
                                            doc.ReplaceText(string.Format("[Pro{0:00}]", i + 1), "");//專案

                                        }


                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("【單身資料】錯誤!");
                            }
                            #endregion

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region //多頁
                        using (DocX doc = DocX.Load(Server.MapPath(secondFilePath)))
                        {
                            wordFileName = string.Format(wordFileName, mocth[0]["TH001"].ToString(), mocth[0]["TH002"].ToString());

                            for (int p = 1; p <= totalPage; p++)
                            {
                                if (p != totalPage)
                                {
                                    #region //預設頁面
                                    if (p != 1)
                                    {
                                        using (DocX subDoc = DocX.Load(Server.MapPath(secondFilePath)))
                                        {
                                            doc.InsertDocument(subDoc, true);
                                        }
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", mocth[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TH001]", mocth[0]["TH001"].ToString());
                                    doc.ReplaceText("[TH002]", mocth[0]["TH002"].ToString());
                                    doc.ReplaceText("[TH005]", mocth[0]["TH005"].ToString());
                                    doc.ReplaceText("[TH006]", mocth[0]["TH006"].ToString());
                                    doc.ReplaceText("[TH029]", mocth[0]["TH029"].ToString());
                                    doc.ReplaceText("[TH003]", mocth[0]["TH003"].ToString());
                                    doc.ReplaceText("[TH004]", mocth[0]["TH004"].ToString());
                                    doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());
                                    doc.ReplaceText("[TH009]", mocth[0]["TH009"].ToString());
                                    doc.ReplaceText("[UDF01]", mocth[0]["UDF01"].ToString());
                                    doc.ReplaceText("[TH032]", mocth[0]["TH032"].ToString());
                                    doc.ReplaceText("[TH018]", mocth[0]["TH018"].ToString());
                                    doc.ReplaceText("[TH019]", mocth[0]["TH019"].ToString());
                                    doc.ReplaceText("[TH021]", mocth[0]["TH021"].ToString());
                                    doc.ReplaceText("[TH022]", mocth[0]["TH022"].ToString());
                                    doc.ReplaceText("[TH027]", mocth[0]["TH027"].ToString());
                                    doc.ReplaceText("[TH020]", mocth[0]["TH020"].ToString());
                                    doc.ReplaceText("[TH031]", mocth[0]["TH031"].ToString());

                                    doc.ReplaceText("[CurrencyTotal]", mocth[0]["CurrencyTotal"].ToString());
                                    doc.ReplaceText("[LocalTotal]", mocth[0]["LocalTotal"].ToString());

                                    if (CurrentCompany == 4)
                                    {
                                        doc.ReplaceText("[CurrencyTotla]", mocth[0]["CurrencyTotal"].ToString());

                                        doc.ReplaceText("[TH007]", mocth[0]["TH007"].ToString());//-
                                        doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());//-
                                        doc.ReplaceText("[TH011]", mocth[0]["TH011"].ToString());//-
                                        doc.ReplaceText("[TH012]", mocth[0]["TH012"].ToString());//-
                                        doc.ReplaceText("[TH013]", mocth[0]["TH013"].ToString());//-
                                        doc.ReplaceText("[TH014]", mocth[0]["TH014"].ToString());//-
                                        doc.ReplaceText("[TH030]", mocth[0]["TH030"].ToString());//-
                                        doc.ReplaceText("[TH015]", mocth[0]["TH015"].ToString());//-
                                        doc.ReplaceText("[TH016]", mocth[0]["TH016"].ToString());//-
                                        doc.ReplaceText("[TH023]", mocth[0]["TH023"].ToString());//-
                                        doc.ReplaceText("[TH033]", mocth[0]["TH033"].ToString());//-
                                    }

                                    doc.ReplaceText("[cp]", p.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[NextPage]", "接下頁...");
                                    doc.ReplaceText("[currendate]", formattedDate);

                                    #endregion

                                    #region //單身
                                    if (mocti.Count() > 0)
                                    {
                                        for (int i = 0; i < 8; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemNo"].ToString());

                                            string tempMtlItemName = mocti[i + (p - 1) * 8]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Quantity"].ToString());
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UomNo"].ToString());
                                            doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Acute"].ToString());

                                            doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i + (p - 1) * 8]["AccertQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i + (p - 1) * 8]["ScrapQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i + (p - 1) * 8]["ReturnQty"]).ToString("F3"));
                                            
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryNo"].ToString());
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryName"].ToString());

                                            doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UnitAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), mocti[i + (p - 1) * 8]["TotalAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["DeduteAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), mocti[i + (p - 1) * 8]["FeeAmount"].ToString());

                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["WoErpFullNo"].ToString());
                                            //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i + (p - 1) * 5]["LotNumber"].ToString());
                                            string tempRemark = mocti[i + (p - 1) * 8]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);

                                            if (CurrentCompany == 4)
                                            {
                                                doc.ReplaceText(string.Format("[InQty{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Quantity"].ToString());
                                                doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());
                                                doc.ReplaceText(string.Format("[DeliverQty{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//已交
                                                doc.ReplaceText(string.Format("[InspectQty{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//待驗
                                                doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//超收
                                                doc.ReplaceText(string.Format("[AmountQry{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//計價
                                                doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//驗收日
                                                doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//批號
                                                doc.ReplaceText(string.Format("[Projrct{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Project"].ToString());//專案
                                                doc.ReplaceText(string.Format("[Pro{0:00}]", i + 1), mocti[i + (p - 1) * 8]["TI015"].ToString());//專案

                                            }

                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion
                                    #endregion
                                }
                                else
                                {
                                    #region //最後一頁
                                    using (DocX finalDoc = DocX.Load(Server.MapPath(filePath)))
                                    {
                                        doc.InsertDocument(finalDoc, true);
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", mocth[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TH001]", mocth[0]["TH001"].ToString());
                                    doc.ReplaceText("[TH002]", mocth[0]["TH002"].ToString());
                                    doc.ReplaceText("[TH005]", mocth[0]["TH005"].ToString());
                                    doc.ReplaceText("[TH006]", mocth[0]["TH006"].ToString());
                                    doc.ReplaceText("[TH029]", mocth[0]["TH029"].ToString());
                                    doc.ReplaceText("[TH003]", mocth[0]["TH003"].ToString());
                                    doc.ReplaceText("[TH004]", mocth[0]["TH004"].ToString());
                                    doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());
                                    doc.ReplaceText("[TH009]", mocth[0]["TH009"].ToString());
                                    doc.ReplaceText("[UDF01]", mocth[0]["UDF01"].ToString());
                                    doc.ReplaceText("[TH032]", mocth[0]["TH032"].ToString());
                                    doc.ReplaceText("[TH018]", mocth[0]["TH018"].ToString());
                                    doc.ReplaceText("[TH019]", mocth[0]["TH019"].ToString());
                                    doc.ReplaceText("[TH021]", mocth[0]["TH021"].ToString());
                                    doc.ReplaceText("[TH022]", mocth[0]["TH022"].ToString());
                                    doc.ReplaceText("[TH027]", mocth[0]["TH027"].ToString());
                                    doc.ReplaceText("[TH020]", mocth[0]["TH020"].ToString());
                                    doc.ReplaceText("[TH031]", mocth[0]["TH031"].ToString());

                                    doc.ReplaceText("[CurrencyTotal]", mocth[0]["CurrencyTotal"].ToString());
                                    doc.ReplaceText("[LocalTotal]", mocth[0]["LocalTotal"].ToString());

                                    if (CurrentCompany == 4)
                                    {
                                        doc.ReplaceText("[CurrencyTotla]", mocth[0]["CurrencyTotal"].ToString());

                                        doc.ReplaceText("[TH007]", mocth[0]["TH007"].ToString());//-
                                        doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());//-
                                        doc.ReplaceText("[TH011]", mocth[0]["TH011"].ToString());//-
                                        doc.ReplaceText("[TH012]", mocth[0]["TH012"].ToString());//-
                                        doc.ReplaceText("[TH013]", mocth[0]["TH013"].ToString());//-
                                        doc.ReplaceText("[TH014]", mocth[0]["TH014"].ToString());//-
                                        doc.ReplaceText("[TH030]", mocth[0]["TH030"].ToString());//-
                                        doc.ReplaceText("[TH015]", mocth[0]["TH015"].ToString());//-
                                        doc.ReplaceText("[TH016]", mocth[0]["TH016"].ToString());//-
                                        doc.ReplaceText("[TH023]", mocth[0]["TH023"].ToString());//-
                                        doc.ReplaceText("[TH033]", mocth[0]["TH033"].ToString());//-
                                    }
                                    doc.ReplaceText("[cp]", totalPage.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[currendate]", formattedDate);

                                    #endregion

                                    #region //單身
                                    if (mocti.Count() > 0)
                                    {
                                        for (int i = 0; i < (mocti.Count() % 8 != 0 ? mocti.Count() % 8 : 8); i++)
                                        {

                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemNo"].ToString());

                                            string tempMtlItemName = mocti[i + (p - 1) * 8]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Quantity"].ToString());
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UomNo"].ToString());
                                            doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Acute"].ToString());

                                            doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AccertQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ScrapQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ReturnQty"]).ToString("F3"));

                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryNo"].ToString());
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryName"].ToString());

                                            doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UnitAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), mocti[i + (p - 1) * 8]["TotalAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["DeduteAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), mocti[i + (p - 1) * 8]["FeeAmount"].ToString());

                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["WoErpFullNo"].ToString());
                                            //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i + (p - 1) * 5]["LotNumber"].ToString());
                                            string tempRemark = mocti[i + (p - 1) * 8]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);

                                            if (CurrentCompany == 4)
                                            {
                                                doc.ReplaceText(string.Format("[InQty{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Quantity"].ToString());
                                                doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());
                                                doc.ReplaceText(string.Format("[DeliverQty{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//已交
                                                doc.ReplaceText(string.Format("[InspectQty{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//待驗
                                                doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//超收
                                                doc.ReplaceText(string.Format("[AmountQry{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//計價
                                                doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//驗收日
                                                doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//批號
                                                doc.ReplaceText(string.Format("[Projrct{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Project"].ToString());//專案
                                                doc.ReplaceText(string.Format("[Pro{0:00}]", i + 1), mocti[i + (p - 1) * 8]["TI015"].ToString());//專案
                                            }

                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion

                                    #region //剩餘欄位
                                    if (mocti.Count() % 8 != 0)
                                    {
                                        for (int i = mocti.Count() % 8; i < 8; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                            if (i == mocti.Count() && i < 8)
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), "");
                                            //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");

                                            if (CurrentCompany == 4)
                                            {
                                                doc.ReplaceText(string.Format("[InQty{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[DeliverQty{0:00}]", i + 1), "");//已交
                                                doc.ReplaceText(string.Format("[InspectQty{0:00}]", i + 1), "");//待驗
                                                doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), "");//超收
                                                doc.ReplaceText(string.Format("[AmountQry{0:00}]", i + 1), "");//計價
                                                doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), "");//驗收日
                                                doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), "");//批號
                                                doc.ReplaceText(string.Format("[Projrct{0:00}]", i + 1), "");//專案
                                                doc.ReplaceText(string.Format("[Pro{0:00}]", i + 1), "");//專案

                                            }

                                        }
                                    }
                                    #endregion
                                    #endregion
                                }
                            }

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }

                    fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = wordFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = wordFileName,
                        fileExtension = ".docx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "目前尚無資料"
                    });
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

        #region //MutieMoWarehouseEnryDocDownload 多張入庫單資料單據下載
        public void MutieMoWarehouseEnryDocDownload(int MweId = -1)
        {
            try
            {
                inventoryDA.AddUserOperateLog(MweId);

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMweDocInfo(MweId);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = "", depositFilePath = "", user = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("單據資料不完整!");

                    #region //判斷公司別文件
                    int CurrentCompany = Convert.ToInt32(Session["UserCompany"]);

                    if (Session["CompanySwitch"] != null)
                    {
                        if (Session["CompanySwitch"].ToString() == "manual")
                        {
                            CurrentCompany = Convert.ToInt32(Session["CompanyId"]);
                        }
                    }

                    var mocth = result["data"];
                    var mocti = result["dataDetail"];

                    switch (CurrentCompany)
                    {
                        case 2: //中揚
                            wordFileName = "入庫單-{0}-{1}";
                            filePath = "~/WordTemplate/MES/P-W001表04-01 入庫單.docx";
                            secondFilePath = "~/WordTemplate/MES/P-W001表04-01 入庫單 換頁.docx";
                            break;
                        case 4: //晶彩
                            wordFileName = "入庫單-{0}-{1}";
                            filePath = "~/WordTemplate/MES/R-PU00-06 入庫單-晶彩.docx";
                            secondFilePath = "~/WordTemplate/MES/R-PU00-06 入庫單-晶彩 換頁.docx";
                            break;
                        default:
                            throw new SystemException("目前公司別未開放功能!");
                    }
                    #endregion

                    #region //產生Doc
                    totalPage = mocti.Count() / 8 + (mocti.Count() % 8 > 0 ? 1 : 0);

                    if (totalPage == 1)
                    {
                        #region //單頁
                        using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                        {
                            wordFileName = string.Format(wordFileName, mocth[0]["TH001"].ToString(), mocth[0]["TH002"].ToString());

                            #region //單頭
                            doc.ReplaceText("[MQ002]", mocth[0]["MQ002"].ToString());//=-
                            doc.ReplaceText("[TH001]", mocth[0]["TH001"].ToString());//=-
                            doc.ReplaceText("[TH002]", mocth[0]["TH002"].ToString());//=-
                            doc.ReplaceText("[TH005]", mocth[0]["TH005"].ToString());//=-
                            doc.ReplaceText("[TH006]", mocth[0]["TH006"].ToString());//=-
                            doc.ReplaceText("[TH029]", mocth[0]["TH029"].ToString());//=-
                            doc.ReplaceText("[TH003]", mocth[0]["TH003"].ToString());//=
                            doc.ReplaceText("[TH004]", mocth[0]["TH004"].ToString());//=-
                            doc.ReplaceText("[TH009]", mocth[0]["TH009"].ToString());//=-
                            doc.ReplaceText("[TH010]", mocth[0]["TH010"].ToString());//=-

                            if (CurrentCompany == 4)
                            {
                                doc.ReplaceText("[TH007]", mocth[0]["TH007"].ToString());//-
                                doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());//-
                                doc.ReplaceText("[TH011]", mocth[0]["TH011"].ToString());//-
                                doc.ReplaceText("[TH012]", mocth[0]["TH012"].ToString());//-
                                doc.ReplaceText("[TH013]", mocth[0]["TH013"].ToString());//-
                                doc.ReplaceText("[TH014]", mocth[0]["TH014"].ToString());//-
                                doc.ReplaceText("[TH030]", mocth[0]["TH030"].ToString());//-
                                doc.ReplaceText("[TH015]", mocth[0]["TH015"].ToString());//-
                                doc.ReplaceText("[TH016]", mocth[0]["TH016"].ToString());//-
                                doc.ReplaceText("[TH023]", mocth[0]["TH023"].ToString());//-
                                doc.ReplaceText("[TH033]", mocth[0]["TH033"].ToString());//-
                            }


                            doc.ReplaceText("[UDF01]", mocth[0]["UDF01"].ToString());//=
                            doc.ReplaceText("[TH032]", mocth[0]["TH032"].ToString());
                            doc.ReplaceText("[TH018]", mocth[0]["TH018"].ToString());
                            doc.ReplaceText("[TH019]", mocth[0]["TH019"].ToString());
                            doc.ReplaceText("[TH021]", mocth[0]["TH021"].ToString());
                            doc.ReplaceText("[TH022]", mocth[0]["TH022"].ToString());
                            doc.ReplaceText("[TH027]", mocth[0]["TH027"].ToString());
                            doc.ReplaceText("[TH020]", mocth[0]["TH020"].ToString());
                            doc.ReplaceText("[TH031]", mocth[0]["TH031"].ToString());

                            doc.ReplaceText("[CurrencyTotal]", mocth[0]["CurrencyTotal"].ToString());
                            doc.ReplaceText("[LocalTotal]", mocth[0]["LocalTotal"].ToString());



                            doc.ReplaceText("[cp]", "1");
                            doc.ReplaceText("[tp]", "1");
                            #endregion

                            #region //單身
                            if (mocti.Count() > 0)
                            {
                                for (int i = 0; i < mocti.Count(); i++)
                                {
                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocti[i]["Seq"].ToString());
                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocti[i]["MtlItemNo"].ToString());
                                    string tempMtlItemName = mocti[i]["MtlItemName"].ToString();
                                    if (tempMtlItemName.Length > 33)
                                    {
                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);
                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocti[i]["MtlItemSpec"].ToString());

                                    doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocti[i]["Quantity"].ToString());
                                    doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocti[i]["UomNo"].ToString());
                                    doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), mocti[i]["Acute"].ToString());

                                    doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AccertQty"]).ToString("F3"));
                                    doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ScrapQty"]).ToString("F3"));
                                    doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ReturnQty"]).ToString("F3"));

                                    doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), mocti[i]["InventoryNo"].ToString());
                                    doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), mocti[i]["InventoryName"].ToString());

                                    doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), mocti[i]["UnitAmount"].ToString());
                                    doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), mocti[i]["TotalAmount"].ToString());
                                    doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), mocti[i]["DeduteAmount"].ToString());
                                    doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), mocti[i]["FeeAmount"].ToString());

                                    if (CurrentCompany == 4)
                                    {
                                        doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), mocti[i]["Verify"].ToString());
                                        doc.ReplaceText(string.Format("[Deliver{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AccertQty"]).ToString("F3"));//已交
                                        doc.ReplaceText(string.Format("[Inspec{0:00}]", i + 1), mocti[i]["Inspec"].ToString());//待驗
                                        doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["OverRate"]).ToString("F3") + '%');//超收
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AmountQty"]).ToString("F3"));//計價
                                        doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), mocti[i]["TI018"].ToString());//驗收日
                                        doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), mocti[i]["LotNo"].ToString());//批號
                                        doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), mocti[i]["Project"].ToString());//專案


                                    }

                                    doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocti[i]["WoErpFullNo"].ToString());
                                    //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());
                                    string tempRemark = mocti[i]["Remark"].ToString();
                                    if (tempRemark.Length > 50)
                                    {
                                        tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                }

                                #region //剩餘欄位
                                if (mocti.Count() < 8)
                                {
                                    for (int i = mocti.Count(); i < 8; i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                        if (i == mocti.Count() && i < 8)
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                        }
                                        else
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                        }
                                        doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), "");
                                        //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());
                                        doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");

                                        if (CurrentCompany == 4)
                                        {
                                            doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Deliver{0:00}]", i + 1), "");//已交
                                            doc.ReplaceText(string.Format("[Inspec{0:00}]", i + 1), "");//待驗
                                            doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), "");//超收
                                            doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");//計價
                                            doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), "");//驗收日
                                            doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), "");//批號
                                            doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), "");//專案


                                        }


                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("【單身資料】錯誤!");
                            }
                            #endregion

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region //多頁
                        using (DocX doc = DocX.Load(Server.MapPath(secondFilePath)))
                        {
                            wordFileName = string.Format(wordFileName, mocth[0]["TH001"].ToString(), mocth[0]["TH002"].ToString());

                            for (int p = 1; p <= totalPage; p++)
                            {
                                if (p != totalPage)
                                {
                                    #region //預設頁面
                                    if (p != 1)
                                    {
                                        using (DocX subDoc = DocX.Load(Server.MapPath(secondFilePath)))
                                        {
                                            doc.InsertDocument(subDoc, true);
                                        }
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", mocth[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TH001]", mocth[0]["TH001"].ToString());
                                    doc.ReplaceText("[TH002]", mocth[0]["TH002"].ToString());
                                    doc.ReplaceText("[TH005]", mocth[0]["TH005"].ToString());
                                    doc.ReplaceText("[TH006]", mocth[0]["TH006"].ToString());
                                    doc.ReplaceText("[TH029]", mocth[0]["TH029"].ToString());
                                    doc.ReplaceText("[TH003]", mocth[0]["TH003"].ToString());
                                    doc.ReplaceText("[TH004]", mocth[0]["TH004"].ToString());
                                    doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());
                                    doc.ReplaceText("[TH009]", mocth[0]["TH009"].ToString());
                                    doc.ReplaceText("[UDF01]", mocth[0]["UDF01"].ToString());
                                    doc.ReplaceText("[TH032]", mocth[0]["TH032"].ToString());
                                    doc.ReplaceText("[TH018]", mocth[0]["TH018"].ToString());
                                    doc.ReplaceText("[TH019]", mocth[0]["TH019"].ToString());
                                    doc.ReplaceText("[TH021]", mocth[0]["TH021"].ToString());
                                    doc.ReplaceText("[TH022]", mocth[0]["TH022"].ToString());
                                    doc.ReplaceText("[TH027]", mocth[0]["TH027"].ToString());
                                    doc.ReplaceText("[TH020]", mocth[0]["TH020"].ToString());
                                    doc.ReplaceText("[TH031]", mocth[0]["TH031"].ToString());

                                    doc.ReplaceText("[CurrencyTotal]", mocth[0]["CurrencyTotal"].ToString());
                                    doc.ReplaceText("[LocalTotal]", mocth[0]["LocalTotal"].ToString());

                                    if (CurrentCompany == 4)
                                    {
                                        doc.ReplaceText("[TH007]", mocth[0]["TH007"].ToString());//-
                                        doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());//-
                                        doc.ReplaceText("[TH011]", mocth[0]["TH011"].ToString());//-
                                        doc.ReplaceText("[TH012]", mocth[0]["TH012"].ToString());//-
                                        doc.ReplaceText("[TH013]", mocth[0]["TH013"].ToString());//-
                                        doc.ReplaceText("[TH014]", mocth[0]["TH014"].ToString());//-
                                        doc.ReplaceText("[TH030]", mocth[0]["TH030"].ToString());//-
                                        doc.ReplaceText("[TH015]", mocth[0]["TH015"].ToString());//-
                                        doc.ReplaceText("[TH016]", mocth[0]["TH016"].ToString());//-
                                        doc.ReplaceText("[TH023]", mocth[0]["TH023"].ToString());//-
                                        doc.ReplaceText("[TH033]", mocth[0]["TH033"].ToString());//-
                                    }

                                    doc.ReplaceText("[cp]", p.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[NextPage]", "接下頁...");
                                    #endregion

                                    #region //單身
                                    if (mocti.Count() > 0)
                                    {
                                        for (int i = 0; i < 8; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemNo"].ToString());

                                            string tempMtlItemName = mocti[i + (p - 1) * 8]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Quantity"].ToString());
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UomNo"].ToString());
                                            doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Acute"].ToString());

                                            doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i + (p - 1) * 8]["AccertQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i + (p - 1) * 8]["ScrapQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i + (p - 1) * 8]["ReturnQty"]).ToString("F3"));

                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryNo"].ToString());
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryName"].ToString());

                                            doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UnitAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), mocti[i + (p - 1) * 8]["TotalAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["DeduteAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), mocti[i + (p - 1) * 8]["FeeAmount"].ToString());

                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["WoErpFullNo"].ToString());
                                            //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i + (p - 1) * 5]["LotNumber"].ToString());
                                            string tempRemark = mocti[i + (p - 1) * 8]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);

                                            if (CurrentCompany == 4)
                                            {
                                                doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());
                                                doc.ReplaceText(string.Format("[Deliver{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//已交
                                                doc.ReplaceText(string.Format("[Inspec{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//待驗
                                                doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//超收
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//計價
                                                doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//驗收日
                                                doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//批號
                                                doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//專案


                                            }

                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion
                                    #endregion
                                }
                                else
                                {
                                    #region //最後一頁
                                    using (DocX finalDoc = DocX.Load(Server.MapPath(filePath)))
                                    {
                                        doc.InsertDocument(finalDoc, true);
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", mocth[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TH001]", mocth[0]["TH001"].ToString());
                                    doc.ReplaceText("[TH002]", mocth[0]["TH002"].ToString());
                                    doc.ReplaceText("[TH005]", mocth[0]["TH005"].ToString());
                                    doc.ReplaceText("[TH006]", mocth[0]["TH006"].ToString());
                                    doc.ReplaceText("[TH029]", mocth[0]["TH029"].ToString());
                                    doc.ReplaceText("[TH003]", mocth[0]["TH003"].ToString());
                                    doc.ReplaceText("[TH004]", mocth[0]["TH004"].ToString());
                                    doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());
                                    doc.ReplaceText("[TH009]", mocth[0]["TH009"].ToString());
                                    doc.ReplaceText("[UDF01]", mocth[0]["UDF01"].ToString());
                                    doc.ReplaceText("[TH032]", mocth[0]["TH032"].ToString());
                                    doc.ReplaceText("[TH018]", mocth[0]["TH018"].ToString());
                                    doc.ReplaceText("[TH019]", mocth[0]["TH019"].ToString());
                                    doc.ReplaceText("[TH021]", mocth[0]["TH021"].ToString());
                                    doc.ReplaceText("[TH022]", mocth[0]["TH022"].ToString());
                                    doc.ReplaceText("[TH027]", mocth[0]["TH027"].ToString());
                                    doc.ReplaceText("[TH020]", mocth[0]["TH020"].ToString());
                                    doc.ReplaceText("[TH031]", mocth[0]["TH031"].ToString());

                                    doc.ReplaceText("[CurrencyTotal]", mocth[0]["CurrencyTotal"].ToString());
                                    doc.ReplaceText("[LocalTotal]", mocth[0]["LocalTotal"].ToString());

                                    if (CurrentCompany == 4)
                                    {
                                        doc.ReplaceText("[TH007]", mocth[0]["TH007"].ToString());//-
                                        doc.ReplaceText("[TH008]", mocth[0]["TH008"].ToString());//-
                                        doc.ReplaceText("[TH011]", mocth[0]["TH011"].ToString());//-
                                        doc.ReplaceText("[TH012]", mocth[0]["TH012"].ToString());//-
                                        doc.ReplaceText("[TH013]", mocth[0]["TH013"].ToString());//-
                                        doc.ReplaceText("[TH014]", mocth[0]["TH014"].ToString());//-
                                        doc.ReplaceText("[TH030]", mocth[0]["TH030"].ToString());//-
                                        doc.ReplaceText("[TH015]", mocth[0]["TH015"].ToString());//-
                                        doc.ReplaceText("[TH016]", mocth[0]["TH016"].ToString());//-
                                        doc.ReplaceText("[TH023]", mocth[0]["TH023"].ToString());//-
                                        doc.ReplaceText("[TH033]", mocth[0]["TH033"].ToString());//-
                                    }
                                    doc.ReplaceText("[cp]", totalPage.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    #endregion

                                    #region //單身
                                    if (mocti.Count() > 0)
                                    {
                                        for (int i = 0; i < (mocti.Count() % 8 != 0 ? mocti.Count() % 8 : 8); i++)
                                        {

                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemNo"].ToString());

                                            string tempMtlItemName = mocti[i + (p - 1) * 8]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), mocti[i + (p - 1) * 8]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Quantity"].ToString());
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UomNo"].ToString());
                                            doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Acute"].ToString());

                                            doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["AccertQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ScrapQty"]).ToString("F3"));
                                            doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), Convert.ToDecimal(mocti[i]["ReturnQty"]).ToString("F3"));

                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryNo"].ToString());
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), mocti[i + (p - 1) * 8]["InventoryName"].ToString());

                                            doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), mocti[i + (p - 1) * 8]["UnitAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), mocti[i + (p - 1) * 8]["TotalAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), mocti[i + (p - 1) * 8]["DeduteAmount"].ToString());
                                            doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), mocti[i + (p - 1) * 8]["FeeAmount"].ToString());

                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["WoErpFullNo"].ToString());
                                            //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i + (p - 1) * 5]["LotNumber"].ToString());
                                            string tempRemark = mocti[i + (p - 1) * 8]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);

                                            if (CurrentCompany == 4)
                                            {
                                                doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());
                                                doc.ReplaceText(string.Format("[Deliver{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//已交
                                                doc.ReplaceText(string.Format("[Inspec{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//待驗
                                                doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//超收
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//計價
                                                doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//驗收日
                                                doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//批號
                                                doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), mocti[i + (p - 1) * 8]["Verify"].ToString());//專案


                                            }

                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion

                                    #region //剩餘欄位
                                    if (mocti.Count() % 8 != 0)
                                    {
                                        for (int i = mocti.Count() % 8; i < 8; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                            if (i == mocti.Count() && i < 8)
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Quantity{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[UomNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Acute{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[AccertQry{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ScrapQry{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ReturnQry{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Unit{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Total{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Dedute{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Fee{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[WoErpFullNo{0:00}]", i + 1), "");
                                            //doc.ReplaceText(string.Format("[LotNumber{0:00}]", i + 1), mocte[i]["LotNumber"].ToString());
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");

                                            if (CurrentCompany == 4)
                                            {
                                                doc.ReplaceText(string.Format("[Verify{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Deliver{0:00}]", i + 1), "");//已交
                                                doc.ReplaceText(string.Format("[Inspec{0:00}]", i + 1), "");//待驗
                                                doc.ReplaceText(string.Format("[OverRate{0:00}]", i + 1), "");//超收
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");//計價
                                                doc.ReplaceText(string.Format("[AcceptD{0:00}]", i + 1), "");//驗收日
                                                doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), "");//批號
                                                doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), "");//專案


                                            }

                                        }
                                    }
                                    #endregion
                                    #endregion
                                }
                            }

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }

                    fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = wordFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = wordFileName,
                        fileExtension = ".docx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "目前尚無資料"
                    });
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

        #region//ExcelMaterialInputStatisticsReport --投料統計報表EXCEL
        public void ExcelMaterialInputStatisticsReport(string WoErpPrefix = "", string WoErpMtlItemName = "", string MtlItemNo = "", string StartDate = "", string EndDate = "", string FilterType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("ExcessReport", "excel");

                #region //Request
                inventoryDA = new InventoryDA();
                dataRequest = inventoryDA.GetMaterialInputStatisticsReport(WoErpPrefix , WoErpMtlItemName, MtlItemNo , StartDate, EndDate, FilterType
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
                    string excelFileName = "【MES2.0】投料統計報表明细 ";
                    string excelsheetName = "投料統計報表明细 ";
                    
                    dynamic[] detailReportDatas = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                   
                    string[] header = new string[] { "制令單號", "產品品号", "產品品名", "產品规格", "材料品号", "材料品名", "ERP應領材料數量", "ERP已領材料數量"
                        , "投料數量", "超領數量", "未使用" };
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

                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.WoErpPrefix.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.WoErpMtlItemNo.ToString().Trim();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.WoErpMtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.WoErpMtlItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.DemandRequisitionQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.RequisitionQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.SUMMatInRegQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.CLQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.WSYQty.ToString();

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

        #endregion
        #endregion
    }
}