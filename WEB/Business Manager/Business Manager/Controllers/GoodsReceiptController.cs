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
using System.Web.Hosting;
using SCMDA;

namespace Business_Manager.Controllers
{
    public class GoodsReceiptController : WebController
    {
        private GoodsReceiptDA goodsReceiptDA = new GoodsReceiptDA();

        #region //View
        public ActionResult GoodsReceiptManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult GrDetail()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetGoodsReceipt 取得進貨單資料
        [HttpPost]
        public void GetGoodsReceipt(int GrId = -1, string GrErpFullNo = "", int SupplierId = -1, int ConfirmUserId = -1, string SearchKey = "", string ConfirmStatus = "", string ClosureStatus = ""
            , string StartDate = "", string EndDate = ""
            , string PoErpFullNo = "", int PoUserId = -1, string PrErpFullNo = "", int PrUserId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "read,constrained-data");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.GetGoodsReceipt(GrId, GrErpFullNo, SupplierId, ConfirmUserId, SearchKey, ConfirmStatus, ClosureStatus
                    , StartDate, EndDate
                    , PoErpFullNo, PoUserId, PrErpFullNo, PrUserId
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

        #region //GetGrDetail 取得進貨單單身資料
        [HttpPost]
        public void GetGrDetail(int GrDetailId = -1, int GrId = -1, string GrErpFullNo = "", string SearchKey = "", int SupplierId = -1, string ConfirmStatus = "", string CloseStatus = "", string GrErpFullNoWithSeq = "", string TransferStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "read,constrained-data");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.GetGrDetail(GrDetailId, GrId, GrErpFullNo, SearchKey, SupplierId, ConfirmStatus, CloseStatus, GrErpFullNoWithSeq, TransferStatus
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

        #region //GetTotal 取得進貨單統整資料
        [HttpPost]
        public void GetTotal(int GrId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "read,constrained-data");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.GetTotal(GrId);
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

        #region //GetDocumentVerification 取得單據性質驗證資料
        [HttpPost]
        public void GetDocumentVerification(int GrId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "read,constrained-data");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.GetDocumentVerification(GrId);
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


        #region //GetLnBarcode 取得批號綁定條碼 GPAI 20240412
        [HttpPost]
        public void GetLnBarcode(string LotNumberNo = "", string LnBarcodeNo = "", int MtlItemId = -1, int LotNumberId = -1, int GrDetailId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "read,constrained-data");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.GetLnBarcode( LotNumberNo,  LnBarcodeNo,  MtlItemId, LotNumberId, GrDetailId
            ,  OrderBy,  PageIndex,  PageSize);
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
        #region //AddGoodsReceipt 新增進貨單資料
        public void AddGoodsReceipt(string GrErpPrefix = "", string DocDate = "", string ReceiptDate = "", int SupplierId = -1, string Remark = ""
            , string CurrencyCode = "", double Exchange = -1, string PaymentTerm = "", int RowCnt = -1
            , string TaxCode = "", string TaxType = "", string InvoiceType = "", double TaxRate = -1, string UiNo = "", string InvoiceDate = ""
            , string InvoiceNo = "", string ApplyYYMM = "", string DeductType = "", string ContactUser = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.AddGoodsReceipt(GrErpPrefix, DocDate, ReceiptDate, SupplierId, Remark
                    , CurrencyCode, Exchange, PaymentTerm, RowCnt
                    , TaxCode, TaxType, InvoiceType, TaxRate, UiNo, InvoiceDate
                    , InvoiceNo, ApplyYYMM, DeductType, ContactUser);
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

        #region //ImportPurchaseOrder 複製前置單據作業(來源:採購單)
        public void ImportPurchaseOrder(int GrId = -1, int PoId = -1, string CopyExchange = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.ImportPurchaseOrder(GrId, PoId, CopyExchange);
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

        #region //ImportPoDetail 從採購單身帶出進貨單身
        public void ImportPoDetail(int GrId = -1, string PoDetailIdList = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.ImportPoDetail(GrId, PoDetailIdList);
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

        #region //AddGrDetail 新增進貨單單身資料
        public void AddGrDetail(int GrId = -1, int PoDetailId = -1, int InventoryId = -1, string AcceptanceDate = "", double ReceiptQty = -1, double ReceiptExpense = -1
            , double AcceptQty = -1, double AvailableQty = -1, double ReturnQty = -1, int UomId = -1, string QcStatus = ""
            , double OrigUnitPrice = -1, double OrigAmount = -1, double OrigDiscountAmt = -1, int MtlItemId = -1
            , string DiscountDescription = "", string Remark = "", string LotNumber = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.AddGrDetail(GrId, PoDetailId, InventoryId, AcceptanceDate, ReceiptQty, ReceiptExpense
                    , AcceptQty, AvailableQty, ReturnQty, UomId, QcStatus
                    , OrigUnitPrice, OrigAmount, OrigDiscountAmt, MtlItemId
                    , DiscountDescription, Remark, LotNumber);
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

        #region //ImportDeliveryOrder 匯入供應商出貨單
        public void ImportDeliveryOrder(int GrId = -1, string DoNo = "", string CopyExchange = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.ImportDeliveryOrder(GrId, DoNo, CopyExchange);
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

        #region //AddLotBindBarcode --新增綁定條碼 
        public void AddLotBindBarcode(int GrDetailId = -1, int BarcodeQty = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.AddLotBindBarcode(GrDetailId, BarcodeQty, BarcodeNo);
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

        #region //AddLotBindCreatBarcode --產出條碼並綁定批號
        public void AddLotBindCreatBarcode(int GrDetailId = -1, int BarcodeQty = -1, string prefix = "", string suffix = "", int seq = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "add");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.AddLotBindCreatBarcode(GrDetailId, BarcodeQty, prefix, suffix, seq);
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

        #region //IntoLotBindBarcode -- 帶入Lot BindBarode Excel
        [HttpPost]

        public void IntoLotBindBarcode(int GrDetailId = -1, int FileId = -1)
        {
            try
            {
                goodsReceiptDA = new GoodsReceiptDA();
                JObject dataRequestJson = new JObject();
                List<LotBarcode> lotBarcodes = new List<LotBarcode>();
                

                    #region //確認檔案資料是否正確
                    List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = goodsReceiptDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);
                string readTextLine = "";

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            //if (item["FileExtension"].ToString().IndexOf("json") != -1)
                            //{
                            //    return lotBarcodes;
                            //}

                            byte[] fileBytes = (byte[])item["FileContent"];
                            readTextLine = Encoding.Unicode.GetString(fileBytes);

                            using (var workbook = new XLWorkbook(new System.IO.MemoryStream(fileBytes)))
                            {
                                var worksheet = workbook.Worksheet(1);

                                // 假设 BarcodeNo 在第一列，COUNT 在第二列
                                foreach (var row in worksheet.RowsUsed())
                                {
                                    string barcodeNo = row.Cell(1).Value.ToString(); // 第一列为 BarcodeNo
                                    int barcodeCount = row.Cell(2).GetValue<int>(); // 第二列为 COUNT

                                   
                                    LotBarcode qcItem = new LotBarcode()
                                    {
                                        BarcodeNo = barcodeNo,
                                        BarcodeCount = barcodeCount

                                    };
                                    lotBarcodes.Add(qcItem);
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

               

                #region //Request

                dataRequest = goodsReceiptDA.IntoLotBindBarcode(lotBarcodes, GrDetailId, FileId);
                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
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

        #region //Update
        #region //UpdateGoodsReceiptManualSynchronize 進貨單資料手動同步
        [HttpPost]
        public void UpdateGoodsReceiptManualSynchronize(string GrErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "sync");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGoodsReceiptManualSynchronize(GrErpFullNo, SyncStartDate, SyncEndDate
                    , NormalSync, TranSync);
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

        #region //UpdateGoodsReceipt 更新進貨單資料
        [HttpPost]
        public void UpdateGoodsReceipt(int GrId = -1, string GrErpPrefix = "", string DocDate = "", string ReceiptDate = "", int SupplierId = -1, string Remark = ""
            , string CurrencyCode = "", double Exchange = -1, string PaymentTerm = "", int RowCnt = -1
            , string TaxCode = "", string TaxType = "", string InvoiceType = "", double TaxRate = -1, string UiNo = "", string InvoiceDate = ""
            , string InvoiceNo = "", string ApplyYYMM = "", string DeductType = "", string ContactUser = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGoodsReceipt(GrId, GrErpPrefix, DocDate, ReceiptDate, SupplierId, Remark
                    , CurrencyCode, Exchange, PaymentTerm, RowCnt
                    , TaxCode, TaxType, InvoiceType, TaxRate, UiNo, InvoiceDate
                    , InvoiceNo, ApplyYYMM, DeductType, ContactUser);
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

        #region //UpdateGrDetail 更新進貨單身資料
        [HttpPost]
        public void UpdateGrDetail(int GrDetailId = -1, int GrId = -1, int PoDetailId = -1, int InventoryId = -1, string AcceptanceDate = "", double ReceiptQty = -1, double ReceiptExpense = -1
            , double AcceptQty = -1, double AvailableQty = -1, double ReturnQty = -1, int UomId = -1, string QcStatus = ""
            , double OrigUnitPrice = -1, double OrigAmount = -1, double OrigDiscountAmt = -1, int MtlItemId = -1
            , string DiscountDescription = "", string Remark = "", string LotNumber = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGrDetail(GrDetailId, GrId, PoDetailId, InventoryId, AcceptanceDate, ReceiptQty, ReceiptExpense
                    , AcceptQty, AvailableQty, ReturnQty, UomId, QcStatus
                    , OrigUnitPrice, OrigAmount, OrigDiscountAmt, MtlItemId
                    , DiscountDescription, Remark, LotNumber);
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

        #region //UpdateTransferGoodsReceipt 拋轉進貨單據
        [HttpPost]
        public void UpdateTransferGoodsReceipt(int GrId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "data-transfer");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateTransferGoodsReceipt(GrId);
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

        #region //UpdateGrLotNumber 更新進貨單身批號資料
        [HttpPost]
        public void UpdateGrLotNumber(int GrDetailId = -1, string LotNumber = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGrLotNumber(GrDetailId, LotNumber);
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

        #region //UpdateReceiptQty 更新進貨數量
        [HttpPost]
        public void UpdateReceiptQty(string UploadJsonString = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateReceiptQty(UploadJsonString);
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

        #region //UpdateGrConfirm 確認進貨單據
        [HttpPost]
        public void UpdateGrConfirm(int GrId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "confirm");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGrConfirm(GrId);
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

        #region //UpdateGrReConfirm 反確認進貨單據
        [HttpPost]
        public void UpdateGrReConfirm(int GrId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "reconfirm");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGrReConfirm(GrId);
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

        #region //UpdateLnBarcode 編輯批號定條碼數量 
        [HttpPost]
        public void UpdateLnBarcode(int LnBarcodeId = -1, int BarcidrQty = -1)
        {
            try
            {
                //WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateLnBarcode(LnBarcodeId, BarcidrQty);
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

        #region //UpdateGrdetailConfirm -- 確認進貨單單筆單身
        [HttpPost]
        public void UpdateGrdetailConfirm(int GrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "reconfirm");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGrdetailConfirm(GrDetailId);
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

        #region //UpdateGrDetailReConfirm -- 反確認進貨單單筆單身
        [HttpPost]
        public void UpdateGrDetailReConfirm(int GrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "reconfirm");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.UpdateGrDetailReConfirm(GrDetailId);
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
        #region //DeleteGoodsReceipt 進貨單資料刪除
        [HttpPost]
        public void DeleteGoodsReceipt(int GrId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "delete");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.DeleteGoodsReceipt(GrId);
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

        #region //DeleteGrDetail 進貨單單身資料刪除
        [HttpPost]
        public void DeleteGrDetail(int GrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "delete");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.DeleteGrDetail(GrDetailId);
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

        #region //BatchDeleteGrDetail -- 批量刪除進貨單詳細資料
        [HttpPost]
        public void BatchDeleteGrDetail(string GrDetailList = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "delete");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.BatchDeleteGrDetail(GrDetailList);
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

        #region //DeleteLnBarcode 刪除批號綁定條碼
        [HttpPost]
        public void DeleteLnBarcode(int LnBarcodeId = -1)
        {
            try
            {
                //WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.DeleteLnBarcode(LnBarcodeId);
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

        #region //DeleteAllBind 刪除批號全部綁定條碼
        [HttpPost]
        public void DeleteAllBind(int GrDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("GoodsReceiptManagement", "update");

                #region //Request
                goodsReceiptDA = new GoodsReceiptDA();
                dataRequest = goodsReceiptDA.DeleteAllBind(GrDetailId);
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
        #region //UpdateGoodsReceiptSynchronize 進貨單資料同步
        [HttpPost]
        [Route("api/ERP/GoodsReceiptSynchronize")]
        public void UpdateGoodsReceiptSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateGoodsReceiptSynchronize");
                #endregion

                #region //Request
                dataRequest = goodsReceiptDA.UpdateGoodsReceiptSynchronize(Company, UpdateDate);
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