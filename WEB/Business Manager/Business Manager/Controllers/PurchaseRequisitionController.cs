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
using System.Threading.Tasks;

namespace Business_Manager.Controllers
{
    public class PurchaseRequisitionController : WebController
    {
        private PurchaseRequisitionDA purchaseRequisitionDA = new PurchaseRequisitionDA();
        private const string SharedFolderErpFile = @"/SharedFolderErpFile";

        #region //View
        public ActionResult PurchaseRequisition()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult PrDetail()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult PrModification()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult HistoryPrice()
        {
            return View();
        }

        public ActionResult IFrameHistoryPrice()
        {
            return View();
        }

        public ActionResult PrConfirmedNotProcuredRecord()
        {
            //
            ViewLoginCheck();
            return View();
        }
        public ActionResult SuggestionsForPurchaseRecord()
        {
            //Rough draft suggestions for purchase
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetPurchaseRequisition 取得請購單資料
        [HttpPost]
        public void GetPurchaseRequisition(int PrId = -1, string PrNo = "", string PrErpPrefix = "", string PrErpNo = "", string Edition = "", string PrStatus = ""
            , string PrDateStartDate = "", string PrDateEndDate = "", string MtlItemNo = "", int DepartmentId = -1, string PrErpFullNo = "", string BpmNo = ""
            , int UserId = -1, string PoErpFullNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPurchaseRequisition(PrId, PrNo, PrErpPrefix, PrErpNo, Edition, PrStatus
                    , PrDateStartDate, PrDateEndDate, MtlItemNo, DepartmentId, PrErpFullNo, BpmNo
                    , UserId, PoErpFullNo
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

        #region //GetPrSequence 取得請購單序號
        [HttpPost]
        public void GetPrSequence(int PrId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPrSequence(PrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrDetail 取得請購單詳細資料
        [HttpPost]
        public async Task GetPrDetail(int PrDetailId = -1,  int PrId = -1, string ConfirmStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = await purchaseRequisitionDA.GetPrDetail(PrDetailId, PrId, ConfirmStatus
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

        #region //GetPrSeq 取得請購單序號
        [HttpPost]
        public void GetPrSeq()
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPrSeq();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrModification 取得請購單資料
        [HttpPost]
        public void GetPrModification(int PrmId = -1, string PrmStatus = "", string PrDateStartDate = "", string PrDateEndDate = "", string ModiDateStartDate = "", string ModiDateEndDate = "", int UserId = -1
            , string PrErpFullNo = "", string PrNo = "", string Edition = "", string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPrModification(PrmId, PrmStatus, PrDateStartDate, PrDateEndDate, ModiDateStartDate, ModiDateEndDate, UserId
                    , PrErpFullNo, PrNo, Edition, MtlItemNo
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

        #region //GetPrmDetail 取得請購單詳細資料
        [HttpPost]
        public void GetPrmDetail(int PrmDetailId = -1, int PrmId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPrmDetail(PrmDetailId, PrmId
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

        #region //GetPrmSequence 取得請購變更單序號
        [HttpPost]
        public void GetPrmSequence(int PrmId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPrmSequence(PrmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlItemTotalUomInfo 取得品號所有可用單位資料
        [HttpPost]
        public void GetMtlItemTotalUomInfo(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetMtlItemTotalUomInfo(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetInventoryQty 取得品號庫存資料
        [HttpPost]
        public void GetInventoryQty(int MtlItemId = -1, int InventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetInventoryQty(MtlItemId, InventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetHistoryPrice 取得特定品號歷史價格
        [HttpPost]
        public void GetHistoryPrice(string MtlItemNo = "")
        {
            try
            {
                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetHistoryPrice(MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetErpLocalCurrency 取得ERP本幣幣別
        [HttpPost]
        public void GetErpLocalCurrency()
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetErpLocalCurrency();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetPrConfirmedNotProcured --寄送請購單已確認但未採購的清單
        public void GetPrConfirmedNotProcured(string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPrConfirmedNotProcured(OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//GetSuggestionsForPurchase --粗胚建議請購清單
        public void GetSuggestionsForPurchase(string OrderBy = "", int PageIndex = -1, int PageSize = -1
            ,string MtlItemNo="",string MB005="",string MB006="",string MB007="",string MB008="",string Condition="")
        {
            try
            {
                //WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetSuggestionsForPurchase(OrderBy, PageIndex, PageSize
                                , MtlItemNo, MB005, MB006, MB007, MB008, Condition);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetPurchaseRequisitionBPM --請購單BPM簽核歷程
        public void GetPurchaseRequisitionBPM(string OrderBy = "", int PageIndex = -1, int PageSize = -1, string PrNo = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "read");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.GetPurchaseRequisitionBPM(OrderBy, PageIndex, PageSize, PrNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetPrModificationBPM --請購變更單BPM簽核歷程
        #endregion

        #endregion

        #region //Add
        #region //AddPurchaseRequisition 新增請購單資料
        [HttpPost]
        public void AddPurchaseRequisition(string PrErpPrefix = "", string DocDate = "", string PrDate = "", string PrRemark = "", string PrFile = "", int UserId = -1, int DepartmentId = -1, string Priority = "", string BomType = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "add");

                string Source = "/" + ControllerContext.RouteData.Values["controller"].ToString() + "/" + ControllerContext.RouteData.Values["action"].ToString();

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.AddPurchaseRequisition(PrErpPrefix, DocDate, PrDate, PrRemark, PrFile, UserId, DepartmentId, Priority, Source, BomType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPrDetail 新增請購單詳細資料
        [HttpPost]
        public void AddPrDetail(int PrId = -1, string PrSequence = "", int MtlItemId = -1, string PrMtlItemName = "", string PrMtlItemSpec = "", int InventoryId = -1, int PrUomId = -1, int PrQty = -1, string DemandDate = ""
            , int SupplierId = -1, string PrCurrency = "", string PrExchangeRate = "", double PrUnitPrice = -1, double PrPrice = -1, double PrPriceTw = -1
            , string UrgentMtl = "", string ProductionPlan = "", string Project = "", int SoDetailId = -1, string PrDetailRemark = "", string PrFile = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "add");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.AddPrDetail(PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate
                    , SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw
                    , UrgentMtl, ProductionPlan, Project, SoDetailId, PrDetailRemark, PrFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPrModification 新增請購變更單資料
        [HttpPost]
        public void AddPrModification(int PrId = -1, int UserId = -1, int DepartmentId = -1, string DocDate = "", string ModiReason = "", string PrmRemark = "", string PrmFile = "", string Priority = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "add");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.AddPrModification(PrId, UserId, DepartmentId, DocDate, ModiReason, PrmRemark, PrmFile, Priority);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPrmDetail 新增請購變更單詳細資料
        [HttpPost]
        public void AddPrmDetail(int PrmId = -1, int PrDetailId = -1, string PrmSequence = "", int MtlItemId = -1, string PrMtlItemName = "", string PrMtlItemSpec = "", int InventoryId = -1, int PrUomId = -1, int PrQty = -1, string DemandDate = ""
            , int SupplierId = -1, string PrCurrency = "", string PrExchangeRate = "", double PrUnitPrice = -1, double PrPrice = -1, double PrPriceTw = -1
            , string UrgentMtl = "", string ProductionPlan = "", string Project = "", int SoDetailId = -1, string PoRemark = "", string PrDetailRemark = "", string ModiReason = "", string ClosureStatus = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "add");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.AddPrmDetail(PrmId, PrDetailId, PrmSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate
                    , SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw
                    , UrgentMtl, ProductionPlan, Project, SoDetailId, PoRemark, PrDetailRemark, ModiReason, ClosureStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPrFile 新增請購單附檔
        [HttpPost]
        public void AddPrFile(int PrId = -1, string PrFile = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "add");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.AddPrFile(PrId, PrFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdatePurchaseRequisition 更新請購單資料
        [HttpPost]
        public void UpdatePurchaseRequisition(int PrId = -1, string PrErpPrefix = "", string DocDate = "", string PrDate = "", string PrRemark = "", string PrFile = "", int UserId = -1, int DepartmentId = -1, string Priority = "", string BomType = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "update");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePurchaseRequisition(PrId, PrErpPrefix, DocDate, PrDate, PrRemark, PrFile, UserId, DepartmentId, Priority, BomType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrDetail 更新請購單詳細資料
        [HttpPost]
        public void UpdatePrDetail(int PrDetailId = -1, int PrId = -1, string PrSequence = "", int MtlItemId = -1, string PrMtlItemName = "", string PrMtlItemSpec = "", int InventoryId = -1, int PrUomId = -1, int PrQty = -1, string DemandDate = ""
            , int SupplierId = -1, string PrCurrency = "", string PrExchangeRate = "", double PrUnitPrice = -1, double PrPrice = -1, double PrPriceTw = -1
            , string UrgentMtl = "", string ProductionPlan = "", string Project = "", int SoDetailId = -1, string PoRemark = "", string PrDetailRemark = "", string PrFile = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "update");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrDetail(PrDetailId, PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate
                    , SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw
                    , UrgentMtl, ProductionPlan, Project, SoDetailId, PoRemark, PrDetailRemark, PrFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrTransferBpm 拋轉請購單據至BPM
        [HttpPost]
        public void UpdatePrTransferBpm(int PrId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "data-transfer");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrTransferBpm(PrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrTransferErp 拋轉請購單據至ERP
        [HttpPost]
        public void UpdatePrTransferErp(int PrId, string BpmNo, string BpmStatus, string ComfirmUser, string ErpFolderRoot, string CompanyNo)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "data-transfer");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                ErpFolderRoot = HostingEnvironment.MapPath("~/" + SharedFolderErpFile);
                dataRequest = purchaseRequisitionDA.UpdatePrTransferErp(PrId, BpmNo, BpmStatus, ComfirmUser, ErpFolderRoot, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrModification 更新請購變更單資料
        [HttpPost]
        public void UpdatePrModification(int PrmId = -1, int PrId = -1, int UserId = -1, int DepartmentId = -1, string DocDate = "", string ModiReason = "", string PrmRemark = "", string PrmFile = "", string Priority = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "update");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrModification(PrmId, PrId, UserId, DepartmentId,  DocDate,  ModiReason,  PrmRemark,  PrmFile, Priority);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrmDetail 更新請購變更單詳細資料
        [HttpPost]
        public void UpdatePrmDetail(int PrmDetailId = -1, int PrDetailId = -1, string PrmSequence = "", int MtlItemId = -1, string PrMtlItemName = "", string PrMtlItemSpec = "", int InventoryId = -1, int PrUomId = -1, int PrQty = -1, string DemandDate = ""
            , int SupplierId = -1, string PrCurrency = "", string PrExchangeRate = "", double PrUnitPrice = -1, double PrPrice = -1, double PrPriceTw = -1
            , string UrgentMtl = "", string ProductionPlan = "", string Project = "", int SoDetailId = -1, string PoRemark = "", string PrDetailRemark = "", string ModiReason = "", string ClosureStatus = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "update");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrmDetail(PrmDetailId, PrDetailId, PrmSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate
                    , SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw
                    , UrgentMtl, ProductionPlan, Project, SoDetailId, PoRemark, PrDetailRemark, ModiReason, ClosureStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrmTransferBpm 拋轉請購變更單據至BPM
        [HttpPost]
        public void UpdatePrmTransferBpm(int PrmId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "data-transfer");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrmTransferBpm(PrmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrmTransferErp 拋轉請購變更單據至ERP
        [HttpPost]
        public void UpdatePrmTransferErp(int PrmId, string BpmNo, string BpmStatus, string ComfirmUser, string ErpFolderRoot, string CompanyNo)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "data-transfer");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                ErpFolderRoot = HostingEnvironment.MapPath("~/" + SharedFolderErpFile);
                dataRequest = purchaseRequisitionDA.UpdatePrmTransferErp(PrmId, BpmNo, BpmStatus, ComfirmUser, ErpFolderRoot, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrVoid 請購單作廢
        [HttpPost]
        public void UpdatePrVoid(int PrId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "void");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrVoid(PrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrmVoid 請購變更單作廢
        [HttpPost]
        public void UpdatePrmVoid(int PrmId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "void");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrmVoid(PrmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrDuplicate 複製請購單
        [HttpPost]
        public void UpdatePrDuplicate(int PrId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "update");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.UpdatePrDuplicate(PrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeletePurchaseRequisition -- 刪除請購單據資料
        [HttpPost]
        public void DeletePurchaseRequisition(int PrId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "delete");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.DeletePurchaseRequisition(PrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePrDetail -- 刪除請購單據詳細資料
        [HttpPost]
        public void DeletePrDetail(int PrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "delete");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.DeletePrDetail(PrDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePrModification -- 刪除請購變更單據資料
        [HttpPost]
        public void DeletePrModification(int PrmId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "delete");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.DeletePrModification(PrmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePrmDetail -- 刪除請購變更單據詳細資料
        [HttpPost]
        public void DeletePrmDetail(int PrmDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "delete");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.DeletePrmDetail(PrmDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //GetBpmPrStatusData -- 取得BPM請購單回傳狀態資料
        [HttpPost]
        public void GetBpmPrStatusData(string id = "", string status = "", string rootId = "", string comfirmUser = "", string bpmNo = "", string company = "")
        {
            try
            {
                string ErpFlag = "P";
                var dataRequestJson = new JObject();

                //先記錄LOG
                dataRequest = purchaseRequisitionDA.AddPrLog(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser);
                dataRequestJson = JObject.Parse(dataRequest);
                int PrLogId = -1;
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    foreach (var item in dataRequestJson["data"])
                    {
                        PrLogId = Convert.ToInt32(item["PrLogId"]);
                    }
                }

                //若狀態為Y，則將請購單據拋轉至ERP
                if (status == "Y")
                {
                    #region //Request
                    string ErpFolderRoot = HostingEnvironment.MapPath("~/" + SharedFolderErpFile);
                    purchaseRequisitionDA = new PurchaseRequisitionDA();
                    dataRequest = purchaseRequisitionDA.UpdatePrTransferErp(Convert.ToInt32(id), bpmNo, status, comfirmUser, ErpFolderRoot, company);
                    #endregion

                    dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequest.IndexOf("errorForDA") != -1)
                    {
                        ErpFlag = "F";

                        #region //將錯誤訊息回寫LOG
                        dataRequest = purchaseRequisitionDA.UpdatePrLogErrorMessage(PrLogId, dataRequestJson["msg"].ToString());
                        JObject logDataRequestJson = JObject.Parse(dataRequest);
                        if (logDataRequestJson["status"].ToString() != "success")
                        {
                            throw new SystemException(logDataRequestJson["msg"].ToString());
                        }
                        #endregion
                    }
                }

                #region //更改MES請購單狀態
                dataRequest = purchaseRequisitionDA.UpdatePrStatus(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser, ErpFlag);
                #endregion

                if (ErpFlag == "F")
                {
                    dataRequest = purchaseRequisitionDA.SendPrErrorMail(Convert.ToInt32(id), "PR");
                    throw new SystemException(dataRequestJson["msg"].ToString());
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

        #region //GetBpmPrmStatusData -- 取得BPM請購變更單回傳狀態資料
        [HttpPost]
        public void GetBpmPrmStatusData(string id = "", string status = "", string rootId = "", string comfirmUser = "", string bpmNo = "", string company = "")
        {
            try
            {
                string ErpFlag = "P";
                var dataRequestJson = new JObject();
                //若狀態為Y，則將請購單據拋轉至ERP
                if (status == "Y" || status == "E")
                {
                    #region //Request
                    purchaseRequisitionDA = new PurchaseRequisitionDA();
                    string ErpFolderRoot = HostingEnvironment.MapPath("~/" + SharedFolderErpFile);
                    dataRequest = purchaseRequisitionDA.UpdatePrmTransferErp(Convert.ToInt32(id), bpmNo, status, comfirmUser, ErpFolderRoot, company);
                    #endregion

                    dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequest.IndexOf("errorForDA") != -1)
                    {
                        ErpFlag = "F";
                    }
                }
                else
                {
                    throw new SystemException("請購變更單拋轉狀態錯誤!(" + status + ")");
                }

                #region //更改MES請購變更單狀態
                dataRequest = purchaseRequisitionDA.UpdatePrmStatus(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser, ErpFlag);

                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequest.IndexOf("errorForDA") != -1)
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                if (ErpFlag == "F")
                {
                    dataRequest = purchaseRequisitionDA.SendPrErrorMail(Convert.ToInt32(id), "PRM");
                    throw new SystemException(dataRequestJson["msg"].ToString());
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
            //Response.Write("123456");
        }
        #endregion

        #region //SendPrErrorMail -- 寄送請購/請購變更異常通知信件
        [HttpPost]
        public void SendPrErrorMail(int DocId, string DocType)
        {
            try
            {
                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.SendPrErrorMail(DocId, DocType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ApiAddPrData -- Api自動開立請購單
        [HttpPost]
        [Route("api/ERP/PurchaseRequisitionData")]
        public void ApiAddPrData(string Company, string SecretKey, string PrInfoJson)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiAddPrData");
                #endregion

                #region //Request
                dataRequest = purchaseRequisitionDA.ApiAddPrData(Company, PrInfoJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ApiAddPrFile -- Api請購單附檔新增
        [HttpPost]
        [Route("api/ERP/PrFile")]
        public void ApiAddPrFile(string Company, string SecretKey, string UserNo, int Id, string Type, string ClientIP)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiAddPrFile");
                #endregion

                #region //傳檔確認
                List<FileModel> files = FileHelper.FileSave(Request.Files);
                #endregion

                #region //Request
                dataRequest = purchaseRequisitionDA.ApiAddPrFile(Company, UserNo, Id, Type, ClientIP, files);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ApiPrTransferBpm -- Api拋轉請購單據至BPM
        [HttpPost]
        [Route("api/ERP/PrTransferBpm")]
        public void ApiPrTransferBpm(string Company, string SecretKey, string UserNo, int PrId)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiPrTransferBpm");
                #endregion

                #region //Request
                dataRequest = purchaseRequisitionDA.ApiPrTransferBpm(Company, UserNo, PrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //For批量開立請購單
        #region //BatchAddPurchaseRequisition 批量開立請購單
        [HttpPost]
        public void BatchAddPurchaseRequisition(string UploadJson = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseRequisition", "update");

                #region //Request
                purchaseRequisitionDA = new PurchaseRequisitionDA();
                dataRequest = purchaseRequisitionDA.BatchAddPurchaseRequisition(UploadJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region// Mail Server
        #region//Mail Send 請購單已確認但未採購的
        [HttpPost]
        [Route("api/ERP/PrConfirmedNotProcuredRecord")]
        public void ApiPrConfirmedNotProcuredRecord(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiPrConfirmedNotProcuredRecord");
                #endregion

                #region //Request
                dataRequest = purchaseRequisitionDA.ApiPrConfirmedNotProcuredRecord(Company);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//Mail Send 粗胚建議請購清單(無條件)
        [HttpPost]
        [Route("api/ERP/SuggestionsForPurchaseRecord")]
        public void ApiSuggestionsForPurchaseRecord(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiSuggestionsForPurchaseRecord");
                #endregion

                #region //Request
                dataRequest = purchaseRequisitionDA.ApiSuggestionsForPurchaseRecord(Company);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//Mail Send 粗胚建議請購清單(有條件)
        [HttpPost]
        [Route("api/ERP/SuggestionsForPurchaseConditionRecord")]
        public void ApiSuggestionsForPurchaseConditionRecord(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "ApiSuggestionsForPurchaseConditionRecord");
                #endregion

                #region //Request
                dataRequest = purchaseRequisitionDA.ApiSuggestionsForPurchaseConditionRecord(Company);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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