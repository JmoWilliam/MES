using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using System;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Drawing;
using System.Threading;
using System.Reflection;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using System.Collections.Generic;
using ZXing;        // for BarcodeWriter
using ZXing.QrCode; // for QRCode Engine
using ZXing.QrCode.Internal;
using ZXing.Common;
using MESDA;
using NiceLabel.SDK;

namespace Business_Manager.Controllers
{
    public class ProductionController : WebController
    {
        private ProductionDA productionDA = new ProductionDA();

        #region //View
        public ActionResult ProdDemandManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult WipOrderManagement()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult TreeStru()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MfgOrderManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MoProcessChangeManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TerminateOfManufactureOrder()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult LetteringChangeManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DepSecondmentManagement()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult WoModificationManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BatchManufactureOrder()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BarcodesScrap()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult OspListManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult OspSupplierManagement()
        {
            return View();
        }
        #endregion

        #region //Get
        #region //GetWipOrder 取得ERP製令資料
        [HttpPost]
        public void GetWipOrder(int WoId = -1, string WoErpPrefix = "", string WoErpFullNo = "", string MtlItemNo = "", string DocDate = ""
            , string StartDate = "", string EndDate = "", string ConfirmStatus = "", string WoStatus = "", string TransferStatus = "", string SearchKey = "", string ErpWoTemporaryList = "",string SoErpNo=""
            , int WoCnt=-1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetWipOrder(WoId, WoErpPrefix, WoErpFullNo, MtlItemNo, DocDate
                    , StartDate, EndDate, ConfirmStatus, WoStatus, TransferStatus, SearchKey, ErpWoTemporaryList, SoErpNo
                    , WoCnt
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

        #region //GetWoDetail 取得ERP製令單身資料
        [HttpPost]
        public void GetWoDetail(int WoId = -1, int WoDetailId = -1, string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = "", string SubstituteStatus =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "woDetail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetWoDetail(WoId, WoDetailId, MtlItemNo, MtlItemName, WoErpFullNo, SubstituteStatus
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

        #region //GetBomSubstitution 取得單身元件的替代料
        [HttpPost]
        public void GetBomSubstitution(int MtlItemId = -1, int BomDetailMtlItemId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "woDetail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBomSubstitution(MtlItemId, BomDetailMtlItemId
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

        #region //GetCheckMtlItemDate 檢查品號生效日失效日
        [HttpPost]
        public void GetCheckMtlItemDate(int MtlItemId = -1, string DocDate = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetCheckMtlItemDate(MtlItemId, DocDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetStorehouseAvailableQty 取得ERP品號的庫存數量
        [HttpPost]
        public void GetStorehouseAvailableQty(string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "woDetail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetStorehouseAvailableQty(MtlItemNo
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

        #region //GetProperty 取得單別性質
        [HttpPost]
        public void GetProperty(string WoErpPrefix = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "woDetail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetProperty(WoErpPrefix);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessor 取得加工廠商資料
        [HttpPost]
        public void GetProcessor(string Processor = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "woDetail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetProcessor(Processor);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCurrencyCode 取得匯率
        [HttpPost]
        public void GetCurrencyCode(string CurrencyCode = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "woDetail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetCurrencyCode(CurrencyCode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetManufactureOrder 取得MES製令資料
        [HttpPost]
        public void GetManufactureOrder(int MoId = -1, int ModeId = -1, string WoErpPrefix = "", string WoErpNo = "", string MtlItemNo = "", string MtlItemName = "", int Quantity = -1, string Status = "", int WoId = -1, string WoErpFullNo = "", string otherInfo = "", string ProjectNo = "", string WoErpFullNoAndSeq = "", int WoSeq = -1
            , string QcTypeStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetManufactureOrder(MoId, ModeId, WoErpPrefix, WoErpNo, MtlItemNo, MtlItemName, Quantity, Status, WoId, WoErpFullNo, otherInfo, ProjectNo, WoErpFullNoAndSeq, WoSeq
                    , QcTypeStatus
                    , OrderBy, PageIndex, PageSize);
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

        #region //GetWoSeq 取得ERP製令在MES中的數量
        [HttpPost]
        public void GetWoSeq(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetWoSeq(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcess 取得MES製令製程
        [HttpPost]
        public void GetMoProcess(int MoProcessId = -1, int MoId = -1, string ManufactureOrder = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoProcess(MoProcessId, MoId, ManufactureOrder
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

        #region //GetMoProcessForOspList 取得MES製令製程(For委外預排製程)
        [HttpPost]
        public void GetMoProcessForOspList(int MoProcessId = -1, int MoId = -1, string WoErpFullNo = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoProcessForOspList(MoProcessId, MoId, WoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoMtlSetting 取得MES製令用料設定
        [HttpPost]
        public void GetMoMtlSetting(int MoMtlSettingId = -1, int MoId = -1, string MtlItemNo = "", string BomSequence = "", int MoProcessId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoMtlSetting(MoMtlSettingId, MoId, MtlItemNo, BomSequence, MoProcessId
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

        #region //GetMoSetting 取得MES製令設定
        [HttpPost]
        public void GetMoSetting(int MoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoSetting(MoId
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

        #region //GetBarcodePrint 取得批量條碼資料
        [HttpPost]
        public void GetBarcodePrint(int PrintId = -1, int MoId = -1, string No = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBarcodePrint(PrintId, MoId, No
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

        #region //GetMoRouting 取得MES製令途程資料
        [HttpPost]
        public void GetMoRouting(int MoId = -1, string RoutingName = "", int ModeId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoRouting(MoId, RoutingName, ModeId
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

        #region //GetLetteringHistory 取得歷程刻字資料
        [HttpPost]
        public void GetLetteringHistory(string WoErpFullNo = "", string Lettering = "", string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetLetteringHistory(WoErpFullNo, Lettering, MtlItemNo
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

        #region //GetBomSequence 取得ERP工單單身序號
        [HttpPost]
        public void GetBomSequence(int WoDetailId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBomSequence(WoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoMtlSettingAttribute 取得MES製令用料屬性檔
        [HttpPost]
        public void GetMoMtlSettingAttribute(int MoMtlSettingAttrId = -1, int MoMtlSettingId = -1, string ItemNo = "", string ItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoMtlSettingAttribute(MoMtlSettingAttrId, MoMtlSettingId, ItemNo, ItemName
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

        #region //GetBindMoProcess 取得物料製令製程資料
        [HttpPost]
        public void GetBindMoProcess(int MoMtlSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBindMoProcess(MoMtlSettingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoItemPart 取得製令刻字號資料
        [HttpPost]
        public void GetMoItemPart(int MoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoItemPart(MoId
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

        #region //GetMoProcessItem 取得製令製程屬性檔
        [HttpPost]
        public void GetMoProcessItem(int MoProcessItemId = -1, int MoProcessId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoProcessItem(MoProcessItemId, MoProcessId
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

        #region //GetProjectNo 取得製令專案代碼資料
        [HttpPost]
        public void GetProjectNo()
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetProjectNo();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcode 取得條碼資訊
        [HttpPost]
        public void GetBarcode(string BarcodeNo = "", int MoId = -1, int MoProcessId = -1, string ItemValue = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBarcode(BarcodeNo, MoId, MoProcessId, ItemValue
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

        #region //GetSumPlanQty 取得所有指定之MES製令預計生產量
        [HttpPost]
        public void GetSumPlanQty(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetSumPlanQty(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcessChange 取得製令途程變更單
        [HttpPost]
        public void GetMoProcessChange(int MpcId = -1, int MoId = -1, string MoNo = "", string StartDocDate = "", string EndDocDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoProcessChange(MpcId, MoId, MoNo, StartDocDate, EndDocDate, Status
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

        #region //GetMpcRoutingProcess 取得途程製程資料
        [HttpPost]
        public void GetMpcRoutingProcess(int MpcRoutingProcessId = -1, int MpcId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMpcRoutingProcess(MpcRoutingProcessId, MpcId
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

        #region //GetMpcBarcode 取得途程變更單綁定條碼
        [HttpPost]
        public void GetMpcBarcode(int MpcId = -1, int MpcBarcodeId = -1, string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMpcBarcode(MpcId, MpcBarcodeId, BarcodeNo
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

        #region //GetCallSlipStatus 取得製令是否有領料
        [HttpPost]
        public void GetCallSlipStatus(int MoId=-1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetCallSlipStatus(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrintImitationLabel 取得模造道具標籤內容資訊查詢
        [HttpPost]
        public void GetPrintImitationLabel(int MoId = -1, string BarcodeNo = "")
        {
            try
            {

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetPrintImitationLabel(MoId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLotBarcodeForLabel 標籤機取得批量條碼或包裝條碼
        [HttpPost]
        [Route("api/MES/GetLotBarcodeForLabel")]
        public void GetLotBarcodeForLabel(string MoId = "", string BarcodeNo = "", string BarcodeType = "", string PrintCnt = "", int IntervalFirst = -1, int IntervalLast = -1)
        {
            try
            {

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetLotBarcodeForLabel(MoId, BarcodeNo, BarcodeType, PrintCnt, IntervalFirst, IntervalLast);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMoRoutingItemProcess 取得製令【刻字站】加工細項
        [HttpPost]
        public void GetMoRoutingItemProcess(int MoId = -1)
        {
            try
            {

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoRoutingItemProcess(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcessAttribute 取得置換歷程
        [HttpPost]
        public void GetBarcodeProcessAttribute(int ProcessAttrLogId = -1, int BarcodeId = -1, string OriWoErpFullNo = "", string NewWoErpFullNo = "", string MtlItemName = "", string BarcodeNo = "", string OriItemValue = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LetteringChangeManagment", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBarcodeProcessAttribute(ProcessAttrLogId, BarcodeId, OriWoErpFullNo, NewWoErpFullNo, MtlItemName, BarcodeNo, OriItemValue
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

        #region //GetBarcodeAttribute 取得刻字屬性置換資訊
        [HttpPost]
        public void GetBarcodeAttribute(int BarcodeAttrId = -1, int BarcodeId = -1, int MoId = -1, string BarcodeNo = "", string ItemNo = "", string ItemValue = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LetteringChangeManagment", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBarcodeAttribute(BarcodeAttrId, BarcodeId, MoId, BarcodeNo, ItemNo, ItemValue
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

        #region //GetBarcodeBindItem
        [HttpPost]
        public void GetBarcodeBindItem(int PrintId = -1)
        {
            try
            {
                //WebApiLoginCheck("GetBarcodeBindItem", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetBarcodeBindItem(PrintId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetInputStatusContent
        [HttpPost]
        public void GetInputStatusContent(string Type = "", string StatusSchema = "")
        {
            try
            {
                //WebApiLoginCheck("GetBarcodeBindItem", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetInputStatusContent(Type, StatusSchema);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDepSecondmentRecord 取得部門借調記錄 -- WuTC 2024-03-18
        [HttpPost]
        public void GetDepSecondmentRecord(string StartDate = "", string EndDate = "", int DsId = -1, int UserId = -1, int DepartmentId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DepSecondmentManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetDepSecondmentRecord(DsId, UserId, DepartmentId, Status, StartDate, EndDate
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

        #region //GetOspList 取得委外預排資料
        [HttpPost]
        public void GetOspList(int OspListId = -1, string WoErpFullNo = "", int DepartmentId = -1, int ProcessId = -1, int SupplierId = -1
            , string MtlItemNo = "", string MtlItemName = "", int CreateUserId = -1, string Status = "", List<int> OspListIds = null, int PoUserId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OspListManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetOspList(OspListId, WoErpFullNo, DepartmentId, ProcessId, SupplierId
                    , MtlItemNo, MtlItemName, CreateUserId, Status, OspListIds, PoUserId
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

        #region //GetProcess 取得製程資料
        [HttpPost]
        public void GetProcess(int ProcessId = -1, int CompanyId = -1, string ProcessNo = "", string ProcessName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetProcess(ProcessId, CompanyId, ProcessNo, ProcessName, Status
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

        #region //GetWoErpFullNo 取得製令單號
        [HttpPost]
        public void GetWoErpFullNo()
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetWoErpFullNo();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetOutsourcingProductionAnonymous 取得托外生產單資料(無權限版本)
        [HttpPost]
        public void GetOutsourcingProductionAnonymous(int OspId)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetOutsourcingProductionAnonymous(OspId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSupplierAnonymous 取得供應商資料(無權限版本)
        [HttpPost]
        public void GetSupplierAnonymous(int OspId)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetSupplierAnonymous(OspId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessCodeAnonymous 取得供應商製程代碼資料(無權限版本)
        [HttpPost]
        public void GetProcessCodeAnonymous(int SupplierId = -1)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetProcessCodeAnonymous(SupplierId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CheckOspDetailSupplierAnonymous 確認託外生產單詳細資料是否已維護好供應商相關資料(無權限版本)
        [HttpPost]
        public void CheckOspDetailSupplierAnonymous(int OspId = -1)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.CheckOspDetailSupplierAnonymous(OspId);
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
        #endregion

        #region //Add
        #region //AddWipOrder 製令單頭資料新增
        [HttpPost]
        public void AddWipOrder(int WoId = -1, string DocDate = "", string WoErpPrefix = "", int MtlItemId = -1, int UomId = -1
            , int? SoDetailId = -1, int InventoryId = -1, int PlanQty = -1, string ExpectedStart = "", string ExpectedEnd = ""
            , int UserId = -1, string WoRemark = "", string Project = "", int ViewCompanyId = -1, string LossRateStatus = ""
            , string Property = "", string ProductionLine = "", string Processor = "", string CurrencyCode = "", double Exchange = -1
            , string TaxCode = "", string TaxType = "", string TaxRate = "", string PricingTerm = "", string PaymentTerm = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddWipOrder(WoId, DocDate, WoErpPrefix, MtlItemId, UomId, SoDetailId, InventoryId
                    , PlanQty, ExpectedStart, ExpectedEnd, UserId, WoRemark, Project, ViewCompanyId, LossRateStatus
                    , Property, ProductionLine, Processor, CurrencyCode, Exchange
                    , TaxCode, TaxType, TaxRate, PricingTerm, PaymentTerm);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddWoDetail 製令單身資料新增
        [HttpPost]
        public void AddWoDetail(int WoId = -1, int MtlItemId = -1, int InventoryId = -1, int UomId = -1, double DemandRequisitionQty = -1
            , string SubstituteStatus = "", string ExpectedRequisitionDate = "", string WipDetailRemark = "", string MaterialProperties = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddWoDetail(WoId, MtlItemId, InventoryId, UomId, DemandRequisitionQty
                    , SubstituteStatus, ExpectedRequisitionDate, WipDetailRemark, MaterialProperties);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBom 套用BOM材料新增
        [HttpPost]
        public void AddBom(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddBom(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBomSubstitution 套用BOM材料新增
        [HttpPost]
        public void AddBomSubstitution(int WoId = -1, int MtlItemId = -1, int PlanQty = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddBomSubstitution(WoId, MtlItemId, PlanQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddManufactureOrder 新增MES製令資料
        [HttpPost]
        public void AddManufactureOrder(int WoId = -1, int WoSeq = -1, int Quantity = -1, int InputQty = -1, int CompleteQty = -1
            , int ScrapQty = -1, int ModeId = -1, string Status = "", string DeliveryProcess = "", string ProjectNo = "",string Remark="")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddManufactureOrder(WoId, WoSeq, Quantity, InputQty, CompleteQty
                    , ScrapQty, Status, ProjectNo,Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoRouting 新增MES製令途程資料
        [HttpPost]
        public void AddMoRouting(string MoRoutingData = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoRouting(MoRoutingData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoMtlSetting 新增MES製令用料設定
        [HttpPost]
        public void AddMoMtlSetting(int MoId = -1, int WoDetailId = -1, int UomId = -1, double CompositionQuantity = -1
            , double Base = -1, double LossRate = -1, string BomSequence = "", string ProcessBinding = "", int MoProcessId = -1
            , string BarcodeCtrl = "", string MainBarcode = "", string ControlType = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoMtlSetting(MoId, WoDetailId, UomId, CompositionQuantity
                    , Base, LossRate, BomSequence, ProcessBinding, MoProcessId, BarcodeCtrl, MainBarcode, ControlType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void AddBarcodePrint(int MoId = -1, int CavityCount = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddBarcodePrint(MoId, CavityCount);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoMtlSettingAttribute 新增MES製令用料屬性檔
        [HttpPost]
        public void AddMoMtlSettingAttribute(int MoMtlSettingId = -1, string ItemNo = "", string ItemName = "", string ItemDesc = "", string ChkUnique = "", string ChkRange = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoMtlSettingAttribute(MoMtlSettingId, ItemNo, ItemName, ItemDesc, ChkUnique, ChkRange);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddEmptyBarcodePrint 新增批量空條碼
        [HttpPost]
        public void AddEmptyBarcodePrint(int AddCount = -1, int MoSettingId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddEmptyBarcodePrint(AddCount, MoSettingId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoItemPart 新增製令刻號
        [HttpPost]
        public void AddMoItemPart(int MoId = -1, string MoItemPartPrefix = "", int Seq = -1, string LetteringType = "", int MoItemPartSortNumberSerial = -1, string MoItemPartSortNumberJump = "", string ParticularNo = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoItemPart(MoId, MoItemPartPrefix, Seq, LetteringType, MoItemPartSortNumberSerial, MoItemPartSortNumberJump, ParticularNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoProcess 新增製令製程資料
        [HttpPost]
        public void AddMoProcess(int MoId = -1, string ProcessData = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoProcess(MoId, ProcessData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoProcessItem 新增製令製程屬性檔
        [HttpPost]
        public void AddMoProcessItem(int MoProcessId = -1, string ItemNo = "", string ChkUnique = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoProcessItem(MoProcessId, ItemNo, ChkUnique);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTerminateOfMfg 新增終止流程Barcode
        [HttpPost]
        public void AddTerminateOfMfg(string BarcodeData = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("TerminateOfManufactureOrder", "end");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddTerminateOfMfg(BarcodeData, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoProcessChange 新增製令途程變更單
        [HttpPost]
        public void AddMoProcessChange(int MoId = -1, int MpcUserId = -1, string DocDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMoProcessChange(MoId, MpcUserId, DocDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMpcRoutingProcessAll 新增途程製程資料
        [HttpPost]
        public void AddMpcRoutingProcessAll(int MpcId = -1, int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMpcRoutingProcessAll(MpcId, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMpcBarcode 新增製令途程變更單綁定條碼
        [HttpPost]
        public void AddMpcBarcode(int MpcId = -1, string BarcodeIds = "")
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddMpcBarcode(MpcId, BarcodeIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBarcodeAttribute 新增屬性置換資料 
        [HttpPost]
        public void AddBarcodeAttribute(int OriBarcodeAttrId = -1, int NewBarcodeAttrId = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("LetteringChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddBarcodeAttribute(OriBarcodeAttrId, NewBarcodeAttrId, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //批量條碼綁定穴號
        [HttpPost]
        public void BarcodePrintBindItem(string BarcodePrintIds = "", int ItemValue = -1, int MoId = -1)
        {
            try
            {
                //WebApiLoginCheck("LetteringChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.BarcodePrintBindItem(BarcodePrintIds, ItemValue, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDepSecondmentRecord-- 新增部門借調記錄 -- WuTC 2024-03-18
        [HttpPost]
        public void AddDepSecondmentRecord(int UserId = -1, int DepartmentId = -1, string StartDate = "", string EndDate = "", string Remark = "")
        {
            try
            {
                if (UserId <= 0) throw new SystemException("【借調人員】不能為空!");
                string Date = DateTime.Now.ToString("yyyy-MM-dd");
                string startdate = Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd");

                int result = DateTime.Compare(Convert.ToDateTime(startdate), Convert.ToDateTime(Date));
                if (result < 0) throw new SystemException("開始日期不得早於當天日期!");

                WebApiLoginCheck("DepSecondmentManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddDepSecondmentRecord(UserId, DepartmentId, StartDate, EndDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddOspList 新增委外預排資料
        [HttpPost]
        public void AddOspList(List<OspListData> UploadJson, string TransferFlag)
        {
            try
            {
                WebApiLoginCheck("OspListManagement", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.AddOspList(UploadJson, TransferFlag);
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
        #endregion

        #region //Update
        #region //UpdateWipOrder 更新MES單頭資料
        [HttpPost]
        public void UpdateWipOrder(int WoId = -1, string DocDate = "", int MtlItemId = -1, int SoDetailId = -1, int UomId = -1, int InventoryId = -1, int PlanQty = -1
            , string ExpectedStart = "", string ExpectedEnd = "", string WoRemark = "", string Project = ""
            , string Property = "", string ProductionLine = "", string Processor = "", string CurrencyCode = "", double Exchange = -1
            , string TaxCode = "", string TaxType = "", string TaxRate = "", string PricingTerm = "", string PaymentTerm = ""
            , int ViewCompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateWipOrder(WoId, DocDate, MtlItemId, SoDetailId, UomId, InventoryId, PlanQty,
                    ExpectedStart, ExpectedEnd, WoRemark, Project
                    , Property, ProductionLine, Processor, CurrencyCode, Exchange
                    , TaxCode, TaxType, TaxRate, PricingTerm, PaymentTerm
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

        #region //UpdateWoDetail 更新MES單身資料
        [HttpPost]
        public void UpdateWoDetail(int WoId = -1, int WoDetailId = -1, int MtlItemId = -1, int InventoryId = -1, int UomId = -1, double DemandRequisitionQty = 0.0
            ,string SubstituteStatus = "", string ExpectedRequisitionDate = "", string WipDetailRemark = "", string MaterialProperties = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateWoDetail(WoId, WoDetailId, MtlItemId, InventoryId, UomId, DemandRequisitionQty
                    ,SubstituteStatus, ExpectedRequisitionDate, WipDetailRemark, MaterialProperties);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWoDetailChangeSubstitution 工單單身替代料替換
        [HttpPost]
        public void UpdateWoDetailChangeSubstitution(int WoDetailId = -1, int SubstitutionId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateWoDetailChangeSubstitution(WoDetailId, SubstitutionId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWipOrderToERP MES工單資料拋轉ERP
        [HttpPost]
        public void UpdateWipOrderToERP(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "confirm");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateWipOrderToERP(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWipOrderReconfirm ERP工單反確認
        [HttpPost]
        public void UpdateWipOrderReconfirm(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "reconfirm");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateWipOrderReconfirm(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWomSynchronize 銷貨單資料手動同步
        //[HttpPost]
        //public void UpdateReceiveOrderSynchronize(string ErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
        //    , string NormalSync = "", string TranSync = "", string CompanyNo = "")
        //{
        //    try
        //    {
        //        WebApiLoginCheck("UpdateWomSynchronize", "sync");

        //        #region //Request
        //        productionDA = new ProductionDA();
        //        dataRequest = productionDA.UpdateWomSynchronize(ErpFullNo, SyncStartDate, SyncEndDate
        //            , NormalSync, TranSync, CompanyNo);
        //        #endregion

        //        #region //Response
        //        jsonResponse = BaseHelper.DAResponse(dataRequest);
        //        #endregion
        //    }
        //    catch (Exception e)
        //    {
        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "error",
        //            msg = e.Message
        //        });
        //        #endregion

        //        logger.Error(e.Message);
        //    }

        //    Response.Write(jsonResponse.ToString());
        //}
        #endregion


        #region //UpdateErpModify (未完成)ERP變更單
        //[HttpPost]
        //public void UpdateErpModify(int WoId = -1, string ModifyDate ="", int SoDetailId =-1, int UomId = -1, int InventoryId = -1, int PlanQty = -1
        //    , string ExpectedStart = "", string ExpectedEnd = "", int UserId = -1, string ModiReason = "", string WoRemark = "", string Project = "")
        //{
        //    try
        //    {
        //        WebApiLoginCheck("WipOrderManagement", "modify");

        //        #region //Request
        //        dataRequest = productionDA.UpdateErpModify(WoId, ModifyDate, SoDetailId, UomId, InventoryId, PlanQty
        //            , ExpectedStart, ExpectedEnd, UserId, ModiReason, WoRemark, Project);
        //        #endregion

        //        #region //Response
        //        jsonResponse = BaseHelper.DAResponse(dataRequest);
        //        #endregion
        //    }
        //    catch (Exception e)
        //    {
        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "error",
        //            msg = e.Message
        //        });
        //        #endregion

        //        logger.Error(e.Message);
        //    }

        //    Response.Write(jsonResponse.ToString());
        //}
        #endregion

        #region //UpdateVoidWipOrder ERP工單作廢
        [HttpPost]
        public void UpdateVoidWipOrder(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "void");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateVoidWipOrder(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateManufactureOrderStatus 更新MES製令啟用狀態
        [HttpPost]
        public void UpdateManufactureOrderStatus(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateManufactureOrderStatus(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoSetting 更新MES製令設定
        [HttpPost]
        public void UpdateMoSetting(int MoId = -1, string LotStatus = "", int LotQty = -1, string MtlControl = "", string NgToBarcode = "", string BarcodePrefix = "", string BarcodePostfix = "", int SequenceLen = -1, int ModeId = -1
            , string DeliveryProcess = "", string Source = "", string PVTQCFlag = "", string OQcCheckType = "", string BarcodeCtrl = "", string TrayBarcode = "", string MrType = "", string OutputBarcodeFlag = "", string NgToDisassembly = ""
            , string VehicleFlag = "", string FullPassQcFlag = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoSetting(MoId, LotStatus, LotQty, MtlControl, NgToBarcode, BarcodePrefix, BarcodePostfix, SequenceLen, ModeId
                    , DeliveryProcess, Source, PVTQCFlag, OQcCheckType, BarcodeCtrl, TrayBarcode, MrType, OutputBarcodeFlag, NgToDisassembly
                    , VehicleFlag, FullPassQcFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoMtlSetting 更新MES製令用料設定
        [HttpPost]
        public void UpdateMoMtlSetting(int MoMtlSettingId = -1, int MoId = -1, int WoDetailId = -1, int UomId = -1, double CompositionQuantity = -1
            , double Base = -1, double LossRate = -1, string BomSequence = "", string ProcessBinding = "", int MoProcessId = -1
            , string BarcodeCtrl = "", string MainBarcode = "", string ControlType = "", string BindType = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoMtlSetting(MoMtlSettingId, MoId, WoDetailId, UomId, CompositionQuantity
                    , Base, LossRate, BomSequence, ProcessBinding, MoProcessId, BarcodeCtrl, MainBarcode, ControlType, BindType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBindType 更新物料上料類型設定
        [HttpPost]
        public void UpdateBindType(string MoMtlSettingInfo = "", string BindType = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBindType(MoMtlSettingInfo, BindType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBindMoProcess 更改物料上料製程
        [HttpPost]
        public void UpdateBindMoProcess(string MoMtlSettingInfo = "", int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBindMoProcess(MoMtlSettingInfo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateManufactureOrder 更新MES製令資料
        [HttpPost]
        public void UpdateManufactureOrder(int MoId = -1, int WoId = -1, int WoSeq = -1, int Quantity = -1,
            int InputQty = -1, int CompleteQty = -1, int ScrapQty = -1, string ProjectNo = "",string Remark="")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateManufactureOrder(MoId, WoId, WoSeq, Quantity, InputQty, CompleteQty, ScrapQty, ProjectNo, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBarcodePrint 更新批量條碼
        [HttpPost]
        public void UpdateBarcodePrint(string BarcodeNo = "", int BarcodeQty = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBarcodePrint(BarcodeNo, BarcodeQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoRoutingSort 更新MES製令途程順序
        [HttpPost]
        public void UpdateMoRoutingSort(string SortNumber)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoRoutingSort(SortNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessBinding 更新物料控管-是否綁定製程狀態
        [HttpPost]
        public void UpdateProcessBinding(int MoMtlSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateProcessBinding(MoMtlSettingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBarcodeCtrl 更新物料控管-是否綁定條碼狀態
        [HttpPost]
        public void UpdateBarcodeCtrl(int MoMtlSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBarcodeCtrl(MoMtlSettingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBarcodeCtrlByLotMode 更新物料控管-是否綁定條碼狀態(批量設定模式)
        [HttpPost]
        public void UpdateBarcodeCtrlByLotMode(string MoMtlSettingIdList = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBarcodeCtrlByLotMode(MoMtlSettingIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMainBarcode 更新物料控管-是否檢核條碼狀態
        [HttpPost]
        public void UpdateMainBarcode(int MoMtlSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMainBarcode(MoMtlSettingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMainBarcodeByLotMode 更新物料控管-是否檢核條碼狀態(批量設定模式)
        [HttpPost]
        public void UpdateMainBarcodeByLotMode(string MoMtlSettingIdList = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMainBarcodeByLotMode(MoMtlSettingIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateControlType 更新物料控管-扣帳方式
        [HttpPost]
        public void UpdateControlType(int MoMtlSettingId = -1, string ControlType = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateControlType(MoMtlSettingId, ControlType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDecompositionFlag 更新物料發料一發多狀況
        [HttpPost]
        public void UpdateDecompositionFlag(int MoMtlSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateDecompositionFlag(MoMtlSettingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDecompositionFlagByLotMode 更新物料發料一發多狀況(批量設定模式)
        [HttpPost]
        public void UpdateDecompositionFlagByLotMode(string MoMtlSettingIdList)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateDecompositionFlagByLotMode(MoMtlSettingIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoMtlSettingAttribute 更新MES製令物料控管屬性檔
        [HttpPost]
        public void UpdateMoMtlSettingAttribute(int MoMtlSettingAttrId = -1, int MoMtlSettingId = -1, string ItemNo = "", string ItemName = "", string ItemDesc = "", string ChkUnique = "", string ChkRange = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoMtlSettingAttribute(MoMtlSettingAttrId, MoMtlSettingId, ItemNo, ItemName, ItemDesc, ChkUnique, ChkRange);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoMtlSettingAttributeChkUnique 更新物料控管屬性檔-檢核重複值
        [HttpPost]
        public void UpdateMoMtlSettingAttributeChkUnique(int MoMtlSettingAttrId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoMtlSettingAttributeChkUnique(MoMtlSettingAttrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoProcessSort 更新製令製程排序
        [HttpPost]
        public void UpdateMoProcessSort(string MoProcessSort = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoProcessSort(MoProcessSort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoProcess 更新製令製程資料
        [HttpPost]
        public void UpdateMoProcess(int MoProcessId = -1, int ProcessId = -1, string ProcessAlias = ""
            , string DisplayStatus = "", string NecessityStatus = "", string ProcessCheckStatus = "", string ProcessCheckType = "", string PackageFlag = ""
            ,string RoutingItemProcessDesc="")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoProcess(MoProcessId, ProcessId, ProcessAlias
                    , DisplayStatus, NecessityStatus, ProcessCheckStatus, ProcessCheckType, PackageFlag, RoutingItemProcessDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoNecessityStatus 更新製令製程必要過站狀態
        [HttpPost]
        public void UpdateMoNecessityStatus(int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoNecessityStatus(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoDisplayStatus 更新製令製程是否顯示在流程卡狀態
        [HttpPost]
        public void UpdateMoDisplayStatus(int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoDisplayStatus(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoProcessItem 更新製令製程屬性檔
        [HttpPost]
        public void UpdateMoProcessItem(int MoProcessItemId = -1, int MoProcessId = -1, string ItemNo = "", string ChkUnique = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMoProcessItem(MoProcessItemId, MoProcessId, ItemNo, ChkUnique);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMpcIdStatus 更新製令途程變更單狀態
        [HttpPost]
        public void UpdateMpcIdStatus(int MpcId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMpcIdStatus(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMpcRoutingProcessSort 更新途程製程排序
        [HttpPost]
        public void UpdateMpcRoutingProcessSort(int MpcId = -1, string RoutingProcessSort = "")
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateMpcRoutingProcessSort(MpcId, RoutingProcessSort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBarcodeAttribute 更新屬性置換資料
        [HttpPost]
        public void UpdateBarcodeAttribute(int ProcessAttrLogId = -1, int OriBarcodeAttrId = -1, int NewBarcodeAttrId = -1
            , int MoId = -1, string NewItemValue = "", string NewSortNumber = "", string Remark = "", string ChkUnique = "")
        {
            try
            {
                WebApiLoginCheck("LetteringChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBarcodeAttribute(ProcessAttrLogId, OriBarcodeAttrId, NewBarcodeAttrId
                    , MoId, NewItemValue, NewSortNumber, Remark, ChkUnique);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBindItem 修改綁定穴號
        [HttpPost]
        public void UpdateBindItem(int PrintId = -1, int ItemValue = -1)
        {
            try
            {
                //WebApiLoginCheck("LetteringChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateBindItem(PrintId, ItemValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion 

        #region //UpdateSelectBindItem 修改選擇綁定穴號
        [HttpPost]
        public void UpdateSelectBindItem(string PrintIds = "", int ItemValue = -1)
        {
            try
            {
                //WebApiLoginCheck("LetteringChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateSelectBindItem(PrintIds, ItemValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCloseManufactureOrder 指定結案ERP及MES製令
        [HttpPost]
        public void UpdateCloseManufactureOrder(int WoId = -1, string Remark = "", string FinishDate = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "finish");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateCloseManufactureOrder(WoId, Remark, FinishDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCloseManufactureOrderBatch 指定結案ERP及MES製令
        [HttpPost]
        public void UpdateCloseManufactureOrderBatch(string WoIdList = "", string Remark = "", string FinishDate = "")
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "finish");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateCloseManufactureOrderBatch(WoIdList, Remark, FinishDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDepSecondmentRecord-- 更新部門借調記錄 -- WuTC 2024-03-18
        [HttpPost]
        public void UpdateDepSecondmentRecord(int DsId = -1, int UserId = -1, int DepartmentId = -1, string StartDate = "", string EndDate = "", string Status = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("DepSecondmentManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateDepSecondmentRecord(DsId, UserId, DepartmentId, StartDate, EndDate, Status, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDepSecondmentRecordStatus-- 變更部門借調記錄狀態 -- WuTC 2024-03-27
        [HttpPost]
        public void UpdateDepSecondmentRecordStatus(int DsId = -1, int UserId = -1, string StartDate = "", string EndDate = "", string Status = "")
        {
            try
            {
                string enddate = Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd");
                string Date = DateTime.Now.ToString("yyyy-MM-dd");
                int endresult = DateTime.Compare(Convert.ToDateTime(EndDate), Convert.ToDateTime(Date));
                if (endresult < 0) throw new SystemException("該筆記錄已過期，請重新建立借調記錄!");

                WebApiLoginCheck("DepSecondmentManagement", "status-switch");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateDepSecondmentRecordStatus(DsId, UserId, StartDate, EndDate, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTransferOspList 拋轉委外預排製程
        [HttpPost]
        public void UpdateTransferOspList(List<OspListData> UploadData)
        {
            try
            {
                WebApiLoginCheck("OspListManagement", "transfer");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateTransferOspList(UploadData);
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

        #region //UpdateOspList 更新委外預排製程
        [HttpPost]
        public void UpdateOspList(int OspListId = -1, int MoProcessId = -1, string OspDate = ""
            , int DepartmentId = -1, int Quantity = -1, string ExpectedDate = "", string ProcessCheckStatus = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("OspListManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateOspList(OspListId, MoProcessId, OspDate
                    , DepartmentId, Quantity, ExpectedDate, ProcessCheckStatus, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateOspSupplierAnonymous 更新託外生產單供應商資料(無權限版本)
        [HttpPost]
        public void UpdateOspSupplierAnonymous(List<OspDetailData> OspDetailData)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateOspSupplierAnonymous(OspDetailData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateOspListStatus 更新委外預排製程狀態 
        [HttpPost]
        public void UpdateOspListStatus(string OspListIds = "", string Status = "", string PoRejectionReason = "")
        {
            try
            {
                WebApiLoginCheck("OspListManagement", "transfer");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.UpdateOspListStatus(OspListIds, Status, PoRejectionReason);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteWipOrder -- MES工單刪除
        [HttpPost]
        public void DeleteWipOrder(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteWipOrder(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion


        #region //DeleteWoDetail -- MES工單單身刪除
        [HttpPost]
        public void DeleteWoDetail(int WoId = -1, int WoDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteWoDetail(WoId, WoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteWoDetailAll -- MES工單單身刪除
        [HttpPost]
        public void DeleteWoDetailAll(int WoId = -1)
        {
            try
            {
                WebApiLoginCheck("WipOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteWoDetailAll(WoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoRouting -- 刪除MES製令途程
        [HttpPost]
        public void DeleteMoRouting(int MoRoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoRouting(MoRoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoMtlSetting -- 刪除MES製令用料設定
        [HttpPost]
        public void DeleteMoMtlSetting(int MoMtlSettingId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoMtlSetting(MoMtlSettingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteManufactureOrder -- 刪除MES製令用料設定
        [HttpPost]
        public void DeleteManufactureOrder(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteManufactureOrder(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBarcodePrint -- 刪除批量條碼資料
        [HttpPost]
        public void DeleteBarcodePrint(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteBarcodePrint(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoMtlSettingAttribute -- 刪除MES製令物料控管屬性檔
        [HttpPost]
        public void DeleteMoMtlSettingAttribute(int MoMtlSettingAttrId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoMtlSettingAttribute(MoMtlSettingAttrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoItemPart -- 刪除MES製令刻字號資料
        [HttpPost]
        public void DeleteMoItemPart(int MoItemPartId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoItemPart(MoItemPartId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteLotMoItemPart -- 批量刪除MES製令刻字號資料
        [HttpPost]
        public void DeleteLotMoItemPart(string MoItemPartId = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteLotMoItemPart(MoItemPartId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteAllMoItemPart -- 刪除指定MES製令刻字號資料
        [HttpPost]
        public void DeleteAllMoItemPart(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteAllMoItemPart(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoProcess -- 刪除MES製令製程
        [HttpPost]
        public void DeleteMoProcess(int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoProcess(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoProcessItem -- 刪除製令製程屬性檔
        [HttpPost]
        public void DeleteMoProcessItem(int MoProcessItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "delete-mo-process");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoProcessItem(MoProcessItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoProcessChange 刪除製令途程變更單
        [HttpPost]
        public void DeleteMoProcessChange(int MpcId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMoProcessChange(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMpcBarcode 刪除製令途程變更單條碼
        [HttpPost]
        public void DeleteMpcBarcode(int MpcBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMpcBarcode(MpcBarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMpcBarcodeAll 刪除製令途程變更單條碼(全)
        [HttpPost]
        public void DeleteMpcBarcodeAll(int MpcId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMpcBarcodeAll(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMpcRoutingProcessAll 刪除製令途程變更單(全)
        [HttpPost]
        public void DeleteMpcRoutingProcessAll(int MpcId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteMpcRoutingProcessAll(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteSelectedBarcodePrint --刪除已選批量條碼
        [HttpPost]
        public void DeleteSelectedBarcodePrint(string BarcodePrintIds = "", int MoId = -1)
        {
            try
            {
                //WebApiLoginCheck("LetteringChangeManagment", "add");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteSelectedBarcodePrint(BarcodePrintIds, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDepSecondmentRecord --刪除部門借調記錄 -- WuTC 2024-03-27
        [HttpPost]
        public void DeleteDepSecondmentRecord(int DsId = -1)
        {
            try
            {
                WebApiLoginCheck("DepSecondmentManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteDepSecondmentRecord(DsId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteOspList -- 刪除委外預排製程
        [HttpPost]
        public void DeleteOspList(string OspListIds = "")
        {
            try
            {
                WebApiLoginCheck("OspListManagement", "delete");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.DeleteOspList(OspListIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Api
        #region //UpdateWipOrderSynchronize -- 製令資料同步
        [HttpPost]
        [Route("api/ERP/WipOrderSynchronize")]
        public void UpdateWipOrderSynchronize(string Company, string SecretKey, string UpdateDate, string WoErpFullNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateWipOrderSynchronize");
                #endregion

                #region //Request
                dataRequest = productionDA.UpdateWipOrderSynchronize(Company, UpdateDate, WoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        //public string TestAPI2()
        //{
        //    //string apiUrl = "http://192.168.20.39:1415/PmdWebSystem/test";
        //    string apiUrl = "http://192.168.134.32:1415/test";
        //    var result = BaseHelper.PostWebRequest(apiUrl);
        //    return "A";
        //}
        #endregion

        #region //Download
        #region //Excel

        #region //ExcelMoTemporary 製令彙整檔匯出Excel
        public void ExcelMoTemporary(string MesMoTemporaryList = "", int ModeId = -1, string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = "", string ProjectNo = "", string Status = "", int Quantity = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "excel");

                #region //Request
                dataRequest = productionDA.ExcelMoTemporary(MesMoTemporaryList, ModeId, MtlItemNo, MtlItemName, WoErpFullNo, ProjectNo, Status, Quantity);
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
                    string excelFileName = "【MES2.0】製令彙整檔";
                    string excelsheetName = "製令彙整頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "QR", "1D", "製令ID", "製令編號", "單據日期", "產品品號", "產品品名", "MES製令數量", "專案代碼" };
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

                            //QRcode
                            var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                            {
                                Format = BarcodeFormat.QR_CODE,
                                Options = new EncodingOptions //設定大小
                                {
                                    Height = 100,
                                    Width = 100,
                                    PureBarcode = true,
                                    Margin = 0
                                }
                            };
                            //一維條碼
                            //var writer2 = new BarcodeWriter  //dll裡面可以看到屬性
                            //{
                            //    Format = BarcodeFormat.CODE_128,
                            //    Options = new EncodingOptions //設定大小
                            //    {
                            //        Height = 50,
                            //        Width = 100,
                            //        PureBarcode = true,
                            //        Margin = 0
                            //    }
                            //};
                            //產生QRcode
                            var MoId = item.MoId.ToString();
                            string MoNo = item.WoErpPrefix.ToString() + "-" + item.WoErpNo.ToString() + "(" + item.WoSeq.ToString() + ")";
                            string Barcode1DMoId = "*" + item.MoId.ToString() + "*";
                            var img1 = writer1.Write(MoId);
                            string FileName1 = "MoId" + MoId;
                            Bitmap myBitmap1 = new Bitmap(img1);
                            string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/MoId/QR/{0}.bmp", FileName1));
                            myBitmap1.Save(filePath1, ImageFormat.Bmp);

                            //var img2 = writer2.Write(MoId);
                            //string FileName2 = "MoId" + MoId;
                            //Bitmap myBitmap2 = new Bitmap(img2);
                            //string filePath2 = Server.MapPath(string.Format("~/PdfTemplate/MES/MoId/1D/{0}.bmp", FileName2));
                            //myBitmap2.Save(filePath2, ImageFormat.Bmp);

                            #region //圖片
                            var MoIdimagePath = Server.MapPath("~/PdfTemplate/MES/MoId/QR/" + FileName1 + ".bmp").ToString();
                            //var MoIdimagePath2 = Server.MapPath("~/PdfTemplate/MES/MoId/1D/" + FileName2 + ".bmp").ToString();
                            rowIndex++;

                            var MoIdimage = worksheet.AddPicture(MoIdimagePath).MoveTo(worksheet.Cell(rowIndex, 1)).Scale(1);
                            //var MoIdimage2 = worksheet.AddPicture(MoIdimagePath2).MoveTo(worksheet.Cell(rowIndex, 2)).Scale(1);

                            worksheet.Row(rowIndex).Height = 80;
                            //worksheet.Range(rowIndex, 1, rowIndex, 10).Merge().Style = titleStyle;
                            #endregion


                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MoId.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = Barcode1DMoId;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = Barcode1DMoId;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Style.Font.SetFontSize(20);
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Style.Font.SetFontName("Free 3 of 9 Extended");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MoId.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = MoNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.DocDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Quantity.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ProjectNo.ToString();

                        }

                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        worksheet.Column(1).Width = 14;
                        worksheet.Column(2).Width = 16;

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

        #region //Pdf
        #region //GetFlowCardPdf 流程卡
        public void GetFlowCardPdf(int MoId = -1,int CompanyId = -1, string BarcodePage = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - 流程卡資料
                dataRequest = productionDA.GetFlowCardPdf(MoId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #region //Request - 刻字條碼資料
                JObject barcodeLettering = new JObject();
                string dataRequest1 = "";

                dataRequest1 = productionDA.GetBarcodeLetteringPdf(MoId);
                barcodeLettering = BaseHelper.DAResponse(dataRequest1);
                #endregion

                string MoNo = "";
                string LotStatus = "";
                string BarcodeCtrl = "";
                string htmlText = "";
                string CompanyName = "";
                string HeadRemark = "";
                int LotQty = 0;

                var writer = new BarcodeWriter  //dll裡面可以看到屬性
                {
                    Format = BarcodeFormat.CODE_128,
                    Options = new EncodingOptions //設定大小
                    {
                        Height = 50,
                        Width = 200,
                        PureBarcode = false,
                        Margin = 0
                    }
                };
                //產生QRcode
                var img = writer.Write(Convert.ToString(MoId));
                Bitmap myBitmap = new Bitmap(img);
                string FileName = "MoId";
                string filePath = Server.MapPath(string.Format( "~/PdfTemplate/MES/{0}.bmp", FileName));
                myBitmap.Save(filePath, ImageFormat.Bmp);

                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("此製令無法進行列印作業!可能因圖層未設定");


                    #region //html

                    var ModeNo = result[0]["ModeNo"].ToString();

                    if (ModeNo == "JMO-A-011" || ModeNo == "JMO-A-012")
                    {
                        
                    }

                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/FlowCard.html"));

                    MoNo = result[0]["MoNo"].ToString();
                    LotStatus = result[0]["LotStatus"].ToString();
                    BarcodeCtrl = result[0]["BarcodeCtrl"].ToString();
                    LotQty = Convert.ToInt32(result[0]["LotQty"]);
                    if (CompanyId == 2)
                    {
                        CompanyName = "中揚光電股份有限公司";
                        HeadRemark = result[0]["MoRemark"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    }
                    else if (CompanyId == 4)
                    {
                        CompanyName = "晶彩光學有限公司";
                        HeadRemark = result[0]["WoRemark"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                    }
                    else
                    {
                        CompanyName = "中揚光電股份有限公司(錯誤)";
                    }
                    MoNo = result[0]["MoNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string MtlItemNo = result[0]["MtlItemNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string MtlItemName = result[0]["MtlItemName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string MtlItemSpec = result[0]["MtlItemSpec"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string InventoryName = result[0]["InventoryName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string UserName = result[0]["UserName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string RountUserName = result[0]["RountUserName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string Quantity = result[0]["Quantity"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string ExpectedEnd = result[0]["ExpectedEnd"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string PlanQty = result[0]["PlanQty"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string Edition = result[0]["Edition"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                    string CustomerDwgNo = result[0]["CustomerDwgNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                    htmlText = htmlText.Replace("[CompanyName]", CompanyName);
                    htmlText = htmlText.Replace("[Barcode]", Server.MapPath("~/PdfTemplate/MES/MoId.bmp"));
                    htmlText = htmlText.Replace("[MoNo]", MoNo);
                    htmlText = htmlText.Replace("[MtlItemNo]", MtlItemNo);
                    htmlText = htmlText.Replace("[MtlItemName]", MtlItemName);
                    htmlText = htmlText.Replace("[MtlItemSpec]", MtlItemSpec);
                    htmlText = htmlText.Replace("[InventoryName]", InventoryName);
                    htmlText = htmlText.Replace("[UserName]", UserName);
                    htmlText = htmlText.Replace("[RountUserName]", RountUserName);
                    htmlText = htmlText.Replace("[Quantity]", Quantity);
                    htmlText = htmlText.Replace("[ExpectedEnd]", ExpectedEnd);
                    htmlText = htmlText.Replace("[PlanQty]", PlanQty);
                    htmlText = htmlText.Replace("[Edition]", Edition);
                    htmlText = htmlText.Replace("[CustomerDwgNo]", CustomerDwgNo);
                    htmlText = htmlText.Replace("[HeadRemark]", HeadRemark);

                    switch (CompanyName) {
                        case "中揚光電股份有限公司":
                            htmlText = htmlText.Replace("[FlowCardNo]", "P-M002表12-01");
                            break;
                        case "晶彩光學有限公司":
                            htmlText = htmlText.Replace("[FlowCardNo]", "R-MF00-06");
                            break;
                        default:
                            htmlText = htmlText.Replace("[FlowCardNo]", "");
                            break;
                    }
                    



                    if (result[0]["BarcodeCtrl"].ToString()=="Y") {
                        htmlText = htmlText.Replace("[BarcodeCtrl]", "條碼控管");
                    }
                    else if(result[0]["BarcodeCtrl"].ToString() == "N")
                    {
                        htmlText = htmlText.Replace("[BarcodeCtrl]", "數量控管");
                    }

                    if (result[0]["OQcCheckType"].ToString()=="N") {
                        htmlText = htmlText.Replace("[OQcCheckType]", "需出貨檢");
                    }
                    else if (result[0]["OQcCheckType"].ToString() == "Y")
                    {
                        //htmlText = htmlText.Replace("[OQcCheckType]", "無需出貨檢");
                    } 

                    string htmlTemplate = htmlText;                    
                    var pageNum = 15;

                    #region //頁面
                    if (result.Count() > 0)
                    {
                        var detail = result.Count();
                        string htmlDetail = "";
                        int mod = detail % pageNum;
                        int page = detail / pageNum + (mod != 0 ? 1 : 0);


                        for(var i=0;i< result.Count(); i++)
                        {
                            htmlDetail += @"<tr>
                                                    <td style='width:7%;height:30px;'>[i]</td>
                                                    <td style='width:13%;'>[ProcessAlias]
                                                    </td>
                                                    <td style='width:25%;'>[RoutingItemProcessDesc]</td>
                                                    <td style='width:7%;'></td>
                                                    <td style='width:5%;'></td>
                                                    <td style='width:5%;'></td>
                                                    <td style='width:20%;'>[Remark]</td>
                                                    <td style='width:10%;'>[ProcessCheckStatus]</td>
                                                    <td style='width:10%;'></td>
                                                </tr>";

                            string ProcessCheckStatus = "";
                            if (result[i]["ProcessCheckStatus"].ToString() == "Y")
                            {
                                ProcessCheckStatus = "Y";
                            }
                            else if (result[i]["ProcessCheckStatus"].ToString() == "N")
                            {
                                ProcessCheckStatus = "N";
                            }
                            else
                            {
                                ProcessCheckStatus = "資料有問題";
                            }

                            string Remark = result[i]["RoutingRemark"].ToString().Replace("<", " &#60; ");
                            if (Remark == "")
                            {
                                Remark = "";
                            }
                            else
                            {
                                Remark = Remark.Replace("\n", "<br/>");
                            }

                            string RoutingItemProcessDesc = result[i]["RoutingItemProcessDesc"].ToString().Replace("<", " &#60; ");
                            if (RoutingItemProcessDesc == "")
                            {
                                RoutingItemProcessDesc = "";
                            }
                            else
                            {
                                RoutingItemProcessDesc = RoutingItemProcessDesc.Replace("\n", "<br/>");
                            }

                            htmlDetail = htmlDetail.Replace("[i]", (i + 1).ToString());
                            htmlDetail = htmlDetail.Replace("[ProcessAlias]", result[i]["ProcessAlias"].ToString().Replace("<", " &#60; "));
                            htmlDetail = htmlDetail.Replace("[RoutingItemProcessDesc]", RoutingItemProcessDesc);
                            htmlDetail = htmlDetail.Replace("[Remark]", Remark);
                            htmlDetail = htmlDetail.Replace("[ProcessCheckStatus]", ProcessCheckStatus);

                        }
                        htmlText = htmlText.Replace("[Detail]", htmlDetail);
                    }
                    else
                    {
                        string htmlDetail = @"<tr>
                                                <td style='width:30px;height:30px;'></td>
                                                <td style='width:80px;'></td>
                                                <td style='width:160px;'>以下空白 / /</td>
                                                <td style='width:30px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:120px;'></td>
                                                <td style='width:50px;'></td>
                                                <td style='width:50px;'></td>
                                              </tr>";

                        for (int i = 0; i < pageNum -1; i++)
                        {
                            htmlDetail += @"<tr>
                                                <td style='width:30px;height:30px;'></td>
                                                <td style='width:80px;'></td>
                                                <td style='width:160px;'>以下空白 / /</td>
                                                <td style='width:30px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:120px;'></td>
                                                <td style='width:50px;'></td>
                                                <td style='width:50px;'></td>
                                            </tr>";
                        }

                        htmlText = htmlText.Replace("[Detail]", htmlDetail);
                    }
                    #endregion

                    #region //條碼頁面
                    if (BarcodeCtrl == "Y" && BarcodePage == "Y")
                    {
                        if (LotStatus != "Y")
                        {
                            
                            string htmlDetail1 = "";
                            string htmlDetail2 = "<table style='width:100%; page-break-after: always;'>" +
                                                    "<tbody>" +
                                                        "<tr>" +
                                                            "<td colspan='10' style='text-align:center; height:30px;'>";
                            htmlDetail2 += "<h1>" + CompanyName + "</h1>";
                            htmlDetail2 +=                  "</td >" +
                                                        "</tr>" +
                                                        "<tr>" +
                                                            "<td colspan='10' style='text-align:center; height:30px;'>" +
                                                                "<table style='width:100%; table-layout:fixed;'>" +
                                                                    "<tr>" +
                                                                        "<td style='width:20%;'>" +
                                                                        "</td>" +
                                                                        "<td style='width:60%; font-size:20px; color:#005fd5;'><h2>條碼</h2> " +
                                                                        "</td>" +
                                                                        "<td style='width:20%; font-size:14px;' valign='bottom'>[MoNo]" +
                                                                        "</td>" +
                                                                    "</tr>" +
                                                                "</table>" +
                                                            "</td>" +
                                                        "</tr>" +
                                                        "<tr>" +
                                                            "<td colspan='10'>&nbsp;</td>" +
                                                        "</tr>" +
                                                    "</tbody>" +
                                                    "<tbody>" +
                                                        "[htmlDetail2]" +
                                                    "</tbody>" +
                                                    "</table>";
                            var rowNum = 6;
                            string frameWidth = "1000px";
                            string phptpWidth = "100px";
                            string height = "70px";
                            int resultCont = Convert.ToInt32(result[0]["Quantity"]);
                            var mod1 = resultCont % rowNum;
                            List<string> barcodeList = new List<string>();
                            for (var n = 1; n <= resultCont; n++)
                            {

                                var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                                {
                                    Format = BarcodeFormat.QR_CODE,
                                    Options = new EncodingOptions //設定大小
                                    {
                                        Height = 10,
                                        Width = 10,
                                        PureBarcode = true,
                                        Margin = 30
                                    }
                                };
                                //產生QRcode
                                var str_no = string.Format("{0:0000}", n);

                                var img1 = writer1.Write(Convert.ToString(MoId) + Convert.ToString(str_no));
                                string FileName1 = "BarcodeNo" + n;
                                Bitmap myBitmap1 = new Bitmap(img1);
                                string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/Barcode/{0}.bmp", FileName1));
                                myBitmap1.Save(filePath1, ImageFormat.Bmp);
                                barcodeList.Add(Convert.ToString(MoId) + Convert.ToString(str_no));


                                if (n % rowNum == 0)
                                {
                                    htmlDetail1 += @"<tr>";
                                    for(var i = 1; i<= rowNum;i++)
                                    {
                                        htmlDetail1 += @"
                                                        < td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                            <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                                [BarcodeNo]
                                                        </td>
                                                        ";
                                    }
                                     htmlDetail1 += @"</tr>";
                                    for (var i = n - rowNum + 1; i <= n; i++)
                                    {
                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                        int index = htmlDetail1.IndexOf("[BarcodeNo]");
                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail1.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                    }
                                    for (var i = n - rowNum + 1; i <= n; i++)
                                    {
                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                        int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/Barcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                    }
                                    barcodeList.Clear();
                                }
                                if (mod1 > 0)
                                {
                                    if (n == resultCont)
                                    {
                                        htmlDetail1 += @"<tr>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo1]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo2]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo3]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo4]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo5]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height+@"; padding:4px; border: 1px solid black;'>
                                                </td>
                                            </tr>
                                                ";
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo1]", barcodeList[0]);
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                        for (var i = n - (n % rowNum) + 1; i <= n; i++)
                                        {
                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                            int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/Barcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                        }
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNoImg]", "");

                                        barcodeList.Clear();
                                    }
                                }
                            }
                            htmlDetail2 = htmlDetail2.Replace("[htmlDetail2]", htmlDetail1);
                            htmlDetail2 = htmlDetail2.Replace("[MoNo]", result[0]["MoNo"].ToString());
                            htmlText = htmlText.Replace("[BarcodeItem]", htmlDetail2);
                        }
                        else if (LotStatus == "Y")
                        {
                            if (LotQty == 1)
                            {
                                #region //Request - 批量條碼資料
                                JObject barcodePrint = new JObject();
                                string dataRequest2 = "";

                                dataRequest2 = productionDA.GetBarcodePrintPdf(MoId);
                                barcodePrint = BaseHelper.DAResponse(dataRequest2);
                                #endregion

                                #region //批量條碼頁面
                                if (barcodePrint["status"].ToString() == "success")
                                {
                                    if (barcodePrint["result"].ToString() != "[]")
                                    {
                                        var result1 = JObject.Parse(dataRequest2)["data"];
                                        dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest2)["data"].ToString());

                                        string htmlDetail1 = "";
                                        string htmlDetail2 = "<table style='width:100%; page-break-after: always;'>" +
                                                                "<tbody>" +
                                                                    "<tr>" +
                                                                        "<td colspan='10' style='text-align:center; height:30px;'><h1>中揚光電股份有限公司</h1></td>" +
                                                                    "</tr>" +
                                                                    "<tr>" +
                                                                        "<td colspan='10' style='text-align:center; height:30px;'>" +
                                                                            "<table style='width:100%; table-layout:fixed;'>" +
                                                                                "<tr>" +
                                                                                    "<td style='width:20%;'>" +
                                                                                    "</td>" +
                                                                                    "<td style='width:60%; font-size:20px; color:#005fd5;'><h2>批量條碼</h2>" +
                                                                                    "</td>" +
                                                                                    "<td style='width:20%; font-size:14px;' valign='bottom'>[MoNo]" +
                                                                                    "</td>" +
                                                                                "</tr>" +
                                                                            "</table>" +
                                                                        "</td>" +
                                                                    "</tr>" +
                                                                    "<tr>" +
                                                                        "<td colspan='10'>&nbsp;</td>" +
                                                                    "</tr>" +
                                                                "</tbody>" +
                                                                "<tbody>" +
                                                                    "[BarcodeItem]" +
                                                                "</tbody>" +
                                                             "</table>";
                                        htmlDetail2 = htmlDetail2.Replace("[MoNo]", result[0]["MoNo"].ToString());

                                        var num = 1;
                                        var rowNum = 10;
                                        var resultCont = result1.Count();
                                        var mod1 = resultCont % rowNum;
                                        List<string> barcodeList = new List<string>();
                                        foreach (var item in data)
                                        {
                                            var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                                            {
                                                Format = BarcodeFormat.QR_CODE,
                                                Options = new EncodingOptions //設定大小
                                                {
                                                    Height = 48,
                                                    Width = 48,
                                                    PureBarcode = true,
                                                    Margin = 0
                                                }
                                            };
                                            //產生QRcode
                                            var img1 = writer1.Write(Convert.ToString(item.BarcodeNo.ToString()));
                                            string FileName1 = "BarcodeNo" + num;
                                            Bitmap myBitmap1 = new Bitmap(img1);
                                            string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/Barcode/{0}.bmp", FileName1));
                                            myBitmap1.Save(filePath1, ImageFormat.Bmp);
                                            barcodeList.Add(Convert.ToString(item.BarcodeNo.ToString()));


                                            if (num % rowNum == 0)
                                            {
                                                htmlDetail1 += @"<tr>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    </tr>
                                                                ";
                                                for (var i = num - rowNum + 1; i <= num; i++)
                                                {
                                                    StringBuilder sb = new StringBuilder(htmlDetail1);
                                                    int index = htmlDetail1.IndexOf("[BarcodeNo]");
                                                    htmlDetail1 = htmlDetail1.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail1.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                                }
                                                for (var i = num - rowNum + 1; i <= num; i++)
                                                {
                                                    StringBuilder sb = new StringBuilder(htmlDetail1);
                                                    int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                                    htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/Barcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                                }
                                                barcodeList.Clear();
                                            }
                                            if (mod1 > 0)
                                            {
                                                if (num == resultCont)
                                                {
                                                    htmlDetail1 += @"<tr>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo1]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo2]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo3]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo4]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo5]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo6]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo7]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo8]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo9]
                                                                        </td>'
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black;'>
                                                                        </td>
                                                                     </tr>
                                                                        ";
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo1]", barcodeList[0]);
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo6]", barcodeList.Count >= 6 ? barcodeList[5] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo7]", barcodeList.Count >= 7 ? barcodeList[6] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo8]", barcodeList.Count >= 8 ? barcodeList[7] : "");
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNo9]", barcodeList.Count >= 9 ? barcodeList[8] : "");
                                                    for (var i = num - (num % rowNum) + 1; i <= num; i++)
                                                    {
                                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                                        int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/Barcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                                    }
                                                    htmlDetail1 = htmlDetail1.Replace("[BarcodeNoImg]", "");

                                                    barcodeList.Clear();
                                                }
                                            }
                                            num++;
                                        }
                                        htmlDetail2 = htmlDetail2.Replace("[BarcodeItem]", htmlDetail1);
                                        htmlText = htmlText.Replace("[BarcodeLettering]", htmlDetail2);
                                    }
                                    else
                                    {
                                        htmlText = htmlText.Replace("[BarcodeLettering]", "");
                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                #region //Request - 批量條碼資料
                                htmlText = htmlText.Replace("[BarcodeLettering]", "");
                                htmlText = htmlText.Replace("[BarcodeItem]", "");
                                #endregion
                            }
                        }
                    }
                    
                    else
                    {
                        htmlText = htmlText.Replace("[BarcodeItem]", "");
                    }
                    #endregion

                    #region //刻字條碼頁面
                    if (barcodeLettering["status"].ToString() == "success")
                    {
                        if (barcodeLettering["result"].ToString() != "[]")
                        {
                            var result1 = JObject.Parse(dataRequest1)["data"];
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest1)["data"].ToString());

                            string htmlDetail1 = "";
                            string htmlDetail2 = "<table style='width:100%; page-break-after: always;'>" +
                                                    "<tbody>" +
                                                        "<tr>" +
                                                            "<td colspan='10' style='text-align:center; height:30px;'>";
                            htmlDetail2 += "<h1>" + CompanyName + "</h1>";
                            htmlDetail2 +=                  "</td>" +
                                                        "</tr>" +
                                                        "<tr>" +
                                                            "<td colspan='10' style='text-align:center; height:30px;'>" +
                                                                "<table style='width:100%; table-layout:fixed;'>" +
                                                                    "<tr>" +
                                                                        "<td style='width:20%;'>" +
                                                                        "</td>" +
                                                                        "<td style='width:60%; font-size:20px; color:#005fd5;'><h2>刻字條碼</h2>" +
                                                                        "</td>" +
                                                                        "<td style='width:20%; font-size:14px;' valign='bottom'>[MoNo]" +
                                                                        "</td>" +
                                                                    "</tr>" +
                                                                "</table>" +
                                                            "</td>" +
                                                        "</tr>" +
                                                        "<tr>" +
                                                            "<td colspan='10'>&nbsp;</td>" +
                                                        "</tr>" +
                                                    "</tbody>" +
                                                    "<tbody>" +
                                                        "[BarcodeItem]" +
                                                    "</tbody>" +
                                                 "</table>";
                            htmlDetail2 = htmlDetail2.Replace("[MoNo]", result1[0]["MoNo"].ToString());

                            var num = 1;
                            var rowNum = 6;
                            string frameWidth = "1000px";
                            string phptpWidth = "100px";
                            string height = "70px";
                            var resultCont = result1.Count();
                            var mod1 = resultCont % rowNum;
                            List<string> barcodeList = new List<string>();
                            foreach (var item in data)
                            {
                                var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                                {
                                    Format = BarcodeFormat.QR_CODE,
                                    Options = new EncodingOptions //設定大小
                                    {
                                        Height = 10,
                                        Width = 10,
                                        PureBarcode = true,
                                        Margin = 30
                                    }
                                };
                                //產生QRcode
                                var img1 = writer1.Write(Convert.ToString(item.MoItemPartId.ToString()));
                                string FileName1 = "BarcodeNo" + num;
                                Bitmap myBitmap1 = new Bitmap(img1);
                                string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/LetteringBarcode/{0}.bmp", FileName1));
                                myBitmap1.Save(filePath1, ImageFormat.Bmp);
                                barcodeList.Add(Convert.ToString(item.MoItemPartNo.ToString()));


                                if (num % rowNum == 0)
                                {
                                    htmlDetail1 += @"<tr>";
                                    for (var i = 1; i <= rowNum; i++)
                                    {
                                        htmlDetail1 += @"
                                                        < td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                            <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                                [BarcodeNo]
                                                        </td>
                                                        ";
                                    }
                                    htmlDetail1 += @"</tr>";
                                    
                                    for (var i = num - rowNum + 1; i <= num; i++)
                                    {
                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                        int index = htmlDetail1.IndexOf("[BarcodeNo]");
                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail1.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                    }
                                    for (var i = num - rowNum + 1; i <= num; i++)
                                    {
                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                        int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/LetteringBarcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                    }
                                    barcodeList.Clear();
                                }
                                if (mod1 > 0)
                                {
                                    if (num == resultCont)
                                    {
                                        htmlDetail1 += @"<tr>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo1]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo2]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo3]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo4]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo5]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black;'>
                                                </td>
                                            </tr>
                                                ";
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo1]", barcodeList[0]);
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                        for (var i = num - (num % rowNum) + 1; i <= num; i++)
                                        {
                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                            int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/LetteringBarcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                        }
                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNoImg]", "");

                                        barcodeList.Clear();
                                    }
                                }
                                num++;
                            }
                            htmlDetail2 = htmlDetail2.Replace("[BarcodeItem]", htmlDetail1);
                            htmlText = htmlText.Replace("[BarcodeLettering]", htmlDetail2);
                        }
                        else
                        {
                            htmlText = htmlText.Replace("[BarcodeLettering]", "");
                        }
                    }
                    #endregion
                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
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

                #region //製令圖片刪除
                FileInfo file = new FileInfo(filePath);
                file.Delete();
                #endregion

                #region //條碼圖片刪除
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/Barcode"));
                FileInfo[] files1 = di.GetFiles();
                foreach (FileInfo file1 in files1)
                {
                    file1.Delete();
                }
                #endregion

                #region //刻字條碼圖片刪除
                DirectoryInfo di2 = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/LetteringBarcode"));
                FileInfo[] files2 = di2.GetFiles();
                foreach (FileInfo file2 in files2)
                {
                    file2.Delete();
                }
                #endregion

                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = "【MES2.0】流程卡" + MoNo,
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

        #region //GetFlowCardPdfBatch 流程卡(批量)
        public void GetFlowCardPdfBatch(string MoIdList = "", int CompanyId = -1, string BarcodePage = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();
                string htmlTextAll = "";
                string htmlText = "";
                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                foreach (var MoIdStr in MoIdList.Split(','))
                {
                    int MoId = Convert.ToInt32(MoIdStr);
                    #region //Request - 流程卡資料
                    dataRequest = productionDA.GetFlowCardPdf(MoId);
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    #region //Request - 刻字條碼資料
                    JObject barcodeLettering = new JObject();
                    string dataRequest1 = "";

                    dataRequest1 = productionDA.GetBarcodeLetteringPdf(MoId);
                    barcodeLettering = BaseHelper.DAResponse(dataRequest1);
                    #endregion

                    string MoNo = "";
                    string LotStatus = "";
                    string BarcodeCtrl = "";
                    string CompanyName = "";
                    string HeadRemark = "";

                    int LotQty = 0;

                    var writer = new BarcodeWriter  //dll裡面可以看到屬性
                    {
                        Format = BarcodeFormat.CODE_128,
                        Options = new EncodingOptions //設定大小
                        {
                            Height = 50,
                            Width = 200,
                            PureBarcode = false,
                            Margin = 0
                        }
                    };

                    //產生QRcode
                    var img = writer.Write(Convert.ToString(MoId));
                    Bitmap myBitmap = new Bitmap(img);
                    string FileName = "MoId"+ MoId;
                    string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/FlowCard/Order/{0}.bmp", FileName));
                    myBitmap.Save(filePath, ImageFormat.Bmp);

                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        if (result.Count() <= 0) throw new SystemException("此製令無法進行列印作業!可能因圖層未設定");


                        #region //html

                        var ModeNo = result[0]["ModeNo"].ToString();

                        if (ModeNo == "JMO-A-011" || ModeNo == "JMO-A-012")
                        {

                        }

                        htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/FlowCard.html"));

                        MoNo = result[0]["MoNo"].ToString();
                        LotStatus = result[0]["LotStatus"].ToString();
                        BarcodeCtrl = result[0]["BarcodeCtrl"].ToString();
                        LotQty = Convert.ToInt32(result[0]["LotQty"]);
                        if (CompanyId == 2)
                        {
                            CompanyName = "中揚光電股份有限公司";
                            HeadRemark = result[0]["MoRemark"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                        }
                        else if (CompanyId == 4)
                        {
                            CompanyName = "晶彩光學有限公司";
                            HeadRemark = result[0]["WoRemark"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                        }
                        else
                        {
                            CompanyName = "中揚光電股份有限公司(錯誤)";
                        }
                        MoNo = result[0]["MoNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string MtlItemNo = result[0]["MtlItemNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string MtlItemName = result[0]["MtlItemName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string MtlItemSpec = result[0]["MtlItemSpec"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string InventoryName = result[0]["InventoryName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string UserName = result[0]["UserName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string RountUserName = result[0]["RountUserName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string Quantity = result[0]["Quantity"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string ExpectedEnd = result[0]["ExpectedEnd"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string PlanQty = result[0]["PlanQty"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string Edition = result[0]["Edition"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string CustomerDwgNo = result[0]["CustomerDwgNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                        htmlText = htmlText.Replace("[CompanyName]", CompanyName);
                        htmlText = htmlText.Replace("[Barcode]", Server.MapPath("~/PdfTemplate/MES/FlowCard/Order/" + FileName + ".bmp"));
                        htmlText = htmlText.Replace("[MoNo]", MoNo);
                        htmlText = htmlText.Replace("[MtlItemNo]", MtlItemNo);
                        htmlText = htmlText.Replace("[MtlItemName]", MtlItemName);
                        htmlText = htmlText.Replace("[MtlItemSpec]", MtlItemSpec);
                        htmlText = htmlText.Replace("[InventoryName]", InventoryName);
                        htmlText = htmlText.Replace("[UserName]", UserName);
                        htmlText = htmlText.Replace("[RountUserName]", RountUserName);
                        htmlText = htmlText.Replace("[Quantity]", Quantity);
                        htmlText = htmlText.Replace("[ExpectedEnd]", ExpectedEnd);
                        htmlText = htmlText.Replace("[PlanQty]", PlanQty);
                        htmlText = htmlText.Replace("[Edition]", Edition);
                        htmlText = htmlText.Replace("[CustomerDwgNo]", CustomerDwgNo);
                        htmlText = htmlText.Replace("[HeadRemark]", HeadRemark);

                        switch (CompanyName)
                        {
                            case "中揚光電股份有限公司":
                                htmlText = htmlText.Replace("[FlowCardNo]", "P-M002表12-01");
                                break;
                            case "晶彩光學有限公司":
                                htmlText = htmlText.Replace("[FlowCardNo]", "R-MF00-06");
                                break;
                            default:
                                htmlText = htmlText.Replace("[FlowCardNo]", "");
                                break;
                        }




                        if (result[0]["BarcodeCtrl"].ToString() == "Y")
                        {
                            htmlText = htmlText.Replace("[BarcodeCtrl]", "條碼控管");
                        }
                        else if (result[0]["BarcodeCtrl"].ToString() == "N")
                        {
                            htmlText = htmlText.Replace("[BarcodeCtrl]", "數量控管");
                        }

                        if (result[0]["OQcCheckType"].ToString() == "N")
                        {
                            htmlText = htmlText.Replace("[OQcCheckType]", "需出貨檢");
                        }
                        else if (result[0]["OQcCheckType"].ToString() == "Y")
                        {
                            //htmlText = htmlText.Replace("[OQcCheckType]", "無需出貨檢");
                        }

                        string htmlTemplate = htmlText;
                        var pageNum = 15;

                        #region //頁面
                        if (result.Count() > 0)
                        {
                            var detail = result.Count();
                            string htmlDetail = "";
                            int mod = detail % pageNum;
                            int page = detail / pageNum + (mod != 0 ? 1 : 0);


                            for (var i = 0; i < result.Count(); i++)
                            {
                                htmlDetail += @"<tr>
                                                    <td style='width:7%;height:30px;'>[i]</td>
                                                    <td style='width:13%;'>[ProcessAlias]
                                                    </td>
                                                    <td style='width:25%;'>[RoutingItemProcessDesc]</td>
                                                    <td style='width:7%;'></td>
                                                    <td style='width:5%;'></td>
                                                    <td style='width:5%;'></td>
                                                    <td style='width:20%;'>[Remark]</td>
                                                    <td style='width:10%;'>[ProcessCheckStatus]</td>
                                                    <td style='width:10%;'></td>
                                                </tr>";

                                string ProcessCheckStatus = "";
                                if (result[i]["ProcessCheckStatus"].ToString() == "Y")
                                {
                                    ProcessCheckStatus = "Y";
                                }
                                else if (result[i]["ProcessCheckStatus"].ToString() == "N")
                                {
                                    ProcessCheckStatus = "N";
                                }
                                else
                                {
                                    ProcessCheckStatus = "資料有問題";
                                }

                                string Remark = result[i]["RoutingRemark"].ToString().Replace("<", " &#60; ");
                                if (Remark == "")
                                {
                                    Remark = "";
                                }
                                else
                                {
                                    Remark = Remark.Replace("\n", "<br/>");
                                }

                                string RoutingItemProcessDesc = result[i]["RoutingItemProcessDesc"].ToString().Replace("<", " &#60; ");
                                if (RoutingItemProcessDesc == "")
                                {
                                    RoutingItemProcessDesc = "";
                                }
                                else
                                {
                                    RoutingItemProcessDesc = RoutingItemProcessDesc.Replace("\n", "<br/>");
                                }

                                htmlDetail = htmlDetail.Replace("[i]", (i + 1).ToString());
                                htmlDetail = htmlDetail.Replace("[ProcessAlias]", result[i]["ProcessAlias"].ToString().Replace("<", " &#60; "));
                                htmlDetail = htmlDetail.Replace("[RoutingItemProcessDesc]", RoutingItemProcessDesc);
                                htmlDetail = htmlDetail.Replace("[Remark]", Remark);
                                htmlDetail = htmlDetail.Replace("[ProcessCheckStatus]", ProcessCheckStatus);

                            }
                            htmlText = htmlText.Replace("[Detail]", htmlDetail);
                        }
                        else
                        {
                            string htmlDetail = @"<tr>
                                                <td style='width:30px;height:30px;'></td>
                                                <td style='width:80px;'></td>
                                                <td style='width:160px;'>以下空白 / /</td>
                                                <td style='width:30px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:120px;'></td>
                                                <td style='width:50px;'></td>
                                                <td style='width:50px;'></td>
                                              </tr>";

                            for (int i = 0; i < pageNum - 1; i++)
                            {
                                htmlDetail += @"<tr>
                                                <td style='width:30px;height:30px;'></td>
                                                <td style='width:80px;'></td>
                                                <td style='width:160px;'>以下空白 / /</td>
                                                <td style='width:30px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:120px;'></td>
                                                <td style='width:50px;'></td>
                                                <td style='width:50px;'></td>
                                            </tr>";
                            }

                            htmlText = htmlText.Replace("[Detail]", htmlDetail);
                        }
                        #endregion

                        #region //條碼頁面
                        if (BarcodeCtrl == "Y" && BarcodePage == "Y")
                        {
                            if (LotStatus != "Y")
                            {

                                string htmlDetail1 = "";
                                string htmlDetail2 = "<table style='width:100%; page-break-after: always;'>" +
                                                        "<tbody>" +
                                                            "<tr>" +
                                                                "<td colspan='10' style='text-align:center; height:30px;'>";
                                htmlDetail2 += "<h1>" + CompanyName + "</h1>";
                                htmlDetail2 += "</td >" +
                                                            "</tr>" +
                                                            "<tr>" +
                                                                "<td colspan='10' style='text-align:center; height:30px;'>" +
                                                                    "<table style='width:100%; table-layout:fixed;'>" +
                                                                        "<tr>" +
                                                                            "<td style='width:20%;'>" +
                                                                            "</td>" +
                                                                            "<td style='width:60%; font-size:20px; color:#005fd5;'><h2>條碼</h2> " +
                                                                            "</td>" +
                                                                            "<td style='width:20%; font-size:14px;' valign='bottom'>[MoNo]" +
                                                                            "</td>" +
                                                                        "</tr>" +
                                                                    "</table>" +
                                                                "</td>" +
                                                            "</tr>" +
                                                            "<tr>" +
                                                                "<td colspan='10'>&nbsp;</td>" +
                                                            "</tr>" +
                                                        "</tbody>" +
                                                        "<tbody>" +
                                                            "[htmlDetail2]" +
                                                        "</tbody>" +
                                                        "</table>";
                                var rowNum = 6;
                                string frameWidth = "1000px";
                                string phptpWidth = "100px";
                                string height = "70px";
                                int resultCont = Convert.ToInt32(result[0]["Quantity"]);
                                var mod1 = resultCont % rowNum;
                                List<string> barcodeList = new List<string>();
                                for (var n = 1; n <= resultCont; n++)
                                {

                                    var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                                    {
                                        Format = BarcodeFormat.QR_CODE,
                                        Options = new EncodingOptions //設定大小
                                        {
                                            Height = 10,
                                            Width = 10,
                                            PureBarcode = true,
                                            Margin = 30
                                        }
                                    };
                                    //產生QRcode
                                    var str_no = string.Format("{0:0000}", n);

                                    var img1 = writer1.Write(Convert.ToString(MoId) + Convert.ToString(str_no));
                                    string FileName1 = "BarcodeNo" + MoId + n;
                                    Bitmap myBitmap1 = new Bitmap(img1);
                                    string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/FlowCard/Barcode/{0}.bmp", FileName1));
                                    myBitmap1.Save(filePath1, ImageFormat.Bmp);
                                    barcodeList.Add(Convert.ToString(MoId) + Convert.ToString(str_no));


                                    if (n % rowNum == 0)
                                    {
                                        htmlDetail1 += @"<tr>";
                                        for (var i = 1; i <= rowNum; i++)
                                        {
                                            htmlDetail1 += @"
                                                        < td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                            <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                                [BarcodeNo]
                                                        </td>
                                                        ";
                                        }
                                        htmlDetail1 += @"</tr>";
                                        for (var i = n - rowNum + 1; i <= n; i++)
                                        {
                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                            int index = htmlDetail1.IndexOf("[BarcodeNo]");
                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail1.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                        }
                                        for (var i = n - rowNum + 1; i <= n; i++)
                                        {
                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                            int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/FlowCard/Barcode/BarcodeNo" + MoId + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                        }
                                        barcodeList.Clear();
                                    }
                                    if (mod1 > 0)
                                    {
                                        if (n == resultCont)
                                        {
                                            htmlDetail1 += @"<tr>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo1]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo2]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo3]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo4]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo5]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black;'>
                                                </td>
                                            </tr>
                                                ";
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo1]", barcodeList[0]);
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                            for (var i = n - (n % rowNum) + 1; i <= n; i++)
                                            {
                                                StringBuilder sb = new StringBuilder(htmlDetail1);
                                                int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                                htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/FlowCard/Barcode/BarcodeNo" + MoId + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                            }
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNoImg]", "");

                                            barcodeList.Clear();
                                        }
                                    }
                                }
                                htmlDetail2 = htmlDetail2.Replace("[htmlDetail2]", htmlDetail1);
                                htmlDetail2 = htmlDetail2.Replace("[MoNo]", result[0]["MoNo"].ToString());
                                htmlText = htmlText.Replace("[BarcodeItem]", htmlDetail2);
                            }
                            else if (LotStatus == "Y")
                            {
                                if (LotQty == 1)
                                {
                                    #region //Request - 批量條碼資料
                                    JObject barcodePrint = new JObject();
                                    string dataRequest2 = "";

                                    dataRequest2 = productionDA.GetBarcodePrintPdf(MoId);
                                    barcodePrint = BaseHelper.DAResponse(dataRequest2);
                                    #endregion

                                    #region //批量條碼頁面
                                    if (barcodePrint["status"].ToString() == "success")
                                    {
                                        if (barcodePrint["result"].ToString() != "[]")
                                        {
                                            var result1 = JObject.Parse(dataRequest2)["data"];
                                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest2)["data"].ToString());

                                            string htmlDetail1 = "";
                                            string htmlDetail2 = "<table style='width:100%; page-break-after: always;'>" +
                                                                    "<tbody>" +
                                                                        "<tr>" +
                                                                            "<td colspan='10' style='text-align:center; height:30px;'><h1>中揚光電股份有限公司</h1></td>" +
                                                                        "</tr>" +
                                                                        "<tr>" +
                                                                            "<td colspan='10' style='text-align:center; height:30px;'>" +
                                                                                "<table style='width:100%; table-layout:fixed;'>" +
                                                                                    "<tr>" +
                                                                                        "<td style='width:20%;'>" +
                                                                                        "</td>" +
                                                                                        "<td style='width:60%; font-size:20px; color:#005fd5;'><h2>批量條碼</h2>" +
                                                                                        "</td>" +
                                                                                        "<td style='width:20%; font-size:14px;' valign='bottom'>[MoNo]" +
                                                                                        "</td>" +
                                                                                    "</tr>" +
                                                                                "</table>" +
                                                                            "</td>" +
                                                                        "</tr>" +
                                                                        "<tr>" +
                                                                            "<td colspan='10'>&nbsp;</td>" +
                                                                        "</tr>" +
                                                                    "</tbody>" +
                                                                    "<tbody>" +
                                                                        "[BarcodeItem]" +
                                                                    "</tbody>" +
                                                                 "</table>";
                                            htmlDetail2 = htmlDetail2.Replace("[MoNo]", result[0]["MoNo"].ToString());

                                            var num = 1;
                                            var rowNum = 10;
                                            var resultCont = result1.Count();
                                            var mod1 = resultCont % rowNum;
                                            List<string> barcodeList = new List<string>();
                                            foreach (var item in data)
                                            {
                                                var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                                                {
                                                    Format = BarcodeFormat.QR_CODE,
                                                    Options = new EncodingOptions //設定大小
                                                    {
                                                        Height = 48,
                                                        Width = 48,
                                                        PureBarcode = true,
                                                        Margin = 0
                                                    }
                                                };
                                                //產生QRcode
                                                var img1 = writer1.Write(Convert.ToString(item.BarcodeNo.ToString()));
                                                string FileName1 = "BarcodeNo" + MoId + num;
                                                Bitmap myBitmap1 = new Bitmap(img1);
                                                string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/FlowCard/Barcode/{0}.bmp", FileName1));
                                                myBitmap1.Save(filePath1, ImageFormat.Bmp);
                                                barcodeList.Add(Convert.ToString(item.BarcodeNo.ToString()));


                                                if (num % rowNum == 0)
                                                {
                                                    htmlDetail1 += @"<tr>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                        <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                            [BarcodeNo]
                                                                    </td>
                                                                    </tr>
                                                                ";
                                                    for (var i = num - rowNum + 1; i <= num; i++)
                                                    {
                                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                                        int index = htmlDetail1.IndexOf("[BarcodeNo]");
                                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail1.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                                    }
                                                    for (var i = num - rowNum + 1; i <= num; i++)
                                                    {
                                                        StringBuilder sb = new StringBuilder(htmlDetail1);
                                                        int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                                        htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/FlowCard/Barcode/BarcodeNo" + MoId + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                                    }
                                                    barcodeList.Clear();
                                                }
                                                if (mod1 > 0)
                                                {
                                                    if (num == resultCont)
                                                    {
                                                        htmlDetail1 += @"<tr>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo1]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo2]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo3]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo4]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo5]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo6]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo7]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo8]
                                                                        </td>
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                                            <img style='width: 55px;' src='[BarcodeNoImg]' alt=''/>
                                                                                [BarcodeNo9]
                                                                        </td>'
                                                                        <td style='width: 57px;height: 30px; padding:4px; border: 1px solid black;'>
                                                                        </td>
                                                                     </tr>
                                                                        ";
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo1]", barcodeList[0]);
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo6]", barcodeList.Count >= 6 ? barcodeList[5] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo7]", barcodeList.Count >= 7 ? barcodeList[6] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo8]", barcodeList.Count >= 8 ? barcodeList[7] : "");
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNo9]", barcodeList.Count >= 9 ? barcodeList[8] : "");
                                                        for (var i = num - (num % rowNum) + 1; i <= num; i++)
                                                        {
                                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                                            int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/FlowCard/Barcode/BarcodeNo" + MoId + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                                        }
                                                        htmlDetail1 = htmlDetail1.Replace("[BarcodeNoImg]", "");

                                                        barcodeList.Clear();
                                                    }
                                                }
                                                num++;
                                            }
                                            htmlDetail2 = htmlDetail2.Replace("[BarcodeItem]", htmlDetail1);
                                            htmlText = htmlText.Replace("[BarcodeLettering]", htmlDetail2);
                                        }
                                        else
                                        {
                                            htmlText = htmlText.Replace("[BarcodeLettering]", "");
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region //Request - 批量條碼資料
                                    htmlText = htmlText.Replace("[BarcodeLettering]", "");
                                    htmlText = htmlText.Replace("[BarcodeItem]", "");
                                    #endregion
                                }
                            }
                        }

                        else
                        {
                            htmlText = htmlText.Replace("[BarcodeItem]", "");
                        }
                        #endregion

                        #region //刻字條碼頁面
                        if (barcodeLettering["status"].ToString() == "success")
                        {
                            if (barcodeLettering["result"].ToString() != "[]")
                            {
                                var result1 = JObject.Parse(dataRequest1)["data"];
                                dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest1)["data"].ToString());

                                string htmlDetail1 = "";
                                string htmlDetail2 = "<table style='width:100%; page-break-after: always;'>" +
                                                        "<tbody>" +
                                                            "<tr>" +
                                                                "<td colspan='10' style='text-align:center; height:30px;'>";
                                htmlDetail2 += "<h1>" + CompanyName + "</h1>";
                                htmlDetail2 += "</td>" +
                                                            "</tr>" +
                                                            "<tr>" +
                                                                "<td colspan='10' style='text-align:center; height:30px;'>" +
                                                                    "<table style='width:100%; table-layout:fixed;'>" +
                                                                        "<tr>" +
                                                                            "<td style='width:20%;'>" +
                                                                            "</td>" +
                                                                            "<td style='width:60%; font-size:20px; color:#005fd5;'><h2>刻字條碼</h2>" +
                                                                            "</td>" +
                                                                            "<td style='width:20%; font-size:14px;' valign='bottom'>[MoNo]" +
                                                                            "</td>" +
                                                                        "</tr>" +
                                                                    "</table>" +
                                                                "</td>" +
                                                            "</tr>" +
                                                            "<tr>" +
                                                                "<td colspan='10'>&nbsp;</td>" +
                                                            "</tr>" +
                                                        "</tbody>" +
                                                        "<tbody>" +
                                                            "[BarcodeItem]" +
                                                        "</tbody>" +
                                                     "</table>";
                                htmlDetail2 = htmlDetail2.Replace("[MoNo]", result1[0]["MoNo"].ToString());

                                var num = 1;
                                var rowNum = 6;
                                string frameWidth = "1000px";
                                string phptpWidth = "100px";
                                string height = "70px";
                                var resultCont = result1.Count();
                                var mod1 = resultCont % rowNum;
                                List<string> barcodeList = new List<string>();
                                foreach (var item in data)
                                {
                                    var writer1 = new BarcodeWriter  //dll裡面可以看到屬性
                                    {
                                        Format = BarcodeFormat.QR_CODE,
                                        Options = new EncodingOptions //設定大小
                                        {
                                            Height = 10,
                                            Width = 10,
                                            PureBarcode = true,
                                            Margin = 30
                                        }
                                    };
                                    //產生QRcode
                                    var img1 = writer1.Write(Convert.ToString(item.MoItemPartId.ToString()));
                                    string FileName1 = "BarcodeNo" + MoId + num;
                                    Bitmap myBitmap1 = new Bitmap(img1);
                                    string filePath1 = Server.MapPath(string.Format("~/PdfTemplate/MES/FlowCard/LetteringBarcode/{0}.bmp", FileName1));
                                    myBitmap1.Save(filePath1, ImageFormat.Bmp);
                                    barcodeList.Add(Convert.ToString(item.MoItemPartNo.ToString()));


                                    if (num % rowNum == 0)
                                    {
                                        htmlDetail1 += @"<tr>";
                                        for (var i = 1; i <= rowNum; i++)
                                        {
                                            htmlDetail1 += @"
                                                        < td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                            <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                                [BarcodeNo]
                                                        </td>
                                                        ";
                                        }
                                        htmlDetail1 += @"</tr>";

                                        for (var i = num - rowNum + 1; i <= num; i++)
                                        {
                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                            int index = htmlDetail1.IndexOf("[BarcodeNo]");
                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail1.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                        }
                                        for (var i = num - rowNum + 1; i <= num; i++)
                                        {
                                            StringBuilder sb = new StringBuilder(htmlDetail1);
                                            int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                            htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/FlowCard/LetteringBarcode/BarcodeNo" + MoId + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                        }
                                        barcodeList.Clear();
                                    }
                                    if (mod1 > 0)
                                    {
                                        if (num == resultCont)
                                        {
                                            htmlDetail1 += @"<tr>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo1]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo2]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo3]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo4]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo5]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black;'>
                                                </td>
                                            </tr>
                                                ";
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo1]", barcodeList[0]);
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                            for (var i = num - (num % rowNum) + 1; i <= num; i++)
                                            {
                                                StringBuilder sb = new StringBuilder(htmlDetail1);
                                                int index = htmlDetail1.IndexOf("[BarcodeNoImg]");
                                                htmlDetail1 = htmlDetail1.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/FlowCard/LetteringBarcode/BarcodeNo" + MoId + i + ".bmp").ToString()) + htmlDetail1.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                            }
                                            htmlDetail1 = htmlDetail1.Replace("[BarcodeNoImg]", "");

                                            barcodeList.Clear();
                                        }
                                    }
                                    num++;
                                }
                                htmlDetail2 = htmlDetail2.Replace("[BarcodeItem]", htmlDetail1);
                                htmlText = htmlText.Replace("[BarcodeLettering]", htmlDetail2);
                            }
                            else
                            {
                                htmlText = htmlText.Replace("[BarcodeLettering]", "");
                            }
                        }
                        #endregion
                        #endregion
                    }
                    htmlTextAll += htmlText;
                }







                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlTextAll); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
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

                #region //製令圖片刪除
                DirectoryInfo di0 = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/FlowCard/Order"));
                FileInfo[] files0 = di0.GetFiles();
                foreach (FileInfo file1 in files0)
                {
                    file1.Delete();
                }
                #endregion

                #region //條碼圖片刪除
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/FlowCard/Barcode"));
                FileInfo[] files1 = di.GetFiles();
                foreach (FileInfo file1 in files1)
                {
                    file1.Delete();
                }
                #endregion

                #region //刻字條碼圖片刪除
                DirectoryInfo di2 = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/FlowCard/LetteringBarcode"));
                FileInfo[] files2 = di2.GetFiles();
                foreach (FileInfo file2 in files2)
                {
                    file2.Delete();
                }
                #endregion

                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = "【MES2.0】"+ dateNow + "流程卡",
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

        #region //GetCallSlipPdf 領料卡(倉庫)
        public void GetCallSlipPdf(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request
                dataRequest = productionDA.GetCallSlipPdf(MoId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string MoNo = "";
                string htmlText = "";


                #region //產生MoId - QRcode圖片
                var writer = new BarcodeWriter  //dll裡面可以看到屬性
                {
                    Format = BarcodeFormat.CODE_128,
                    Options = new EncodingOptions //設定大小
                    {
                        Height = 50,
                        Width = 200,
                        PureBarcode = false,
                        Margin = 0
                    }
                };
                var img = writer.Write(Convert.ToString(MoId));
                string FileName = "MoId";
                Bitmap myBitmap = new Bitmap(img);
                string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/{0}.bmp", FileName));
                myBitmap.Save(filePath, ImageFormat.Bmp);
                #endregion


                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    if (result.Count() <= 0) throw new SystemException("資料有問題,此製令無法進行列印作業!");

                    #region //取得線別名稱
                    string ProductionLine = "";
                    string ProductionLineName = "";

                    ProductionLine =result[0]["ProductionLine"].ToString();
                    if(ProductionLine != "")
                    {
                        dataRequest = productionDA.GetProductionLine(ProductionLine);
                        jsonResponse = BaseHelper.DAResponse(dataRequest);
                        if (jsonResponse["status"].ToString() == "success")
                        {
                            var result1 = JObject.Parse(dataRequest)["data"];
                            if (result1.Count() <= 0) throw new SystemException("線別資料代碼有問題,找不到");
                            ProductionLineName = result1[0]["ProduceLineName"].ToString();
                            ProductionLine = "(" + result[0]["ProductionLine"].ToString() + ")";
                        }
                    }
                    else
                    {
                        ProductionLine = "";
                    }
                    
                    #endregion


                    #region //html
                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/CallSlip.html"));

                    MoNo = result[0]["MoNo"].ToString();

                    htmlText = htmlText.Replace("[Barcode]", Server.MapPath("~/PdfTemplate/MES/MoId.bmp"));

                    htmlText = htmlText.Replace("[MoNo]", result[0]["MoNo"].ToString());
                    htmlText = htmlText.Replace("[MtlItemNo]", result[0]["MtlItemNo"].ToString());
                    htmlText = htmlText.Replace("[MtlItemName]", result[0]["MtlItemName"].ToString());
                    htmlText = htmlText.Replace("[MtlItemSpec]", result[0]["MtlItemSpec"].ToString());
                    htmlText = htmlText.Replace("[InventoryNo]", result[0]["InventoryNo"].ToString());
                    htmlText = htmlText.Replace("[InventoryName]", result[0]["InventoryName"].ToString());
                    htmlText = htmlText.Replace("[UserName]", result[0]["UserName"].ToString());                    
                    htmlText = htmlText.Replace("[Quantity]", result[0]["Quantity"].ToString());
                    htmlText = htmlText.Replace("[ExpectedEnd]", result[0]["ExpectedEnd"].ToString());
                    htmlText = htmlText.Replace("[PlanQty]", result[0]["PlanQty"].ToString());
                    htmlText = htmlText.Replace("[ProductionLine]", ProductionLine);
                    htmlText = htmlText.Replace("[ProductionLineName]", ProductionLineName);

                    string htmlTemplate = htmlText;

                    var pageNum = 15;

                    if (result.Count() > 0)
                    {
                        var detail = result.Count();
                        string htmlDetail = "";
                        int mod = detail % pageNum;
                        int page = detail / pageNum + (mod != 0 ? 1 : 0);


                        for (var i = 0; i < result.Count(); i++)
                        {
                            #region //取得庫存數量
                            var MtlItemNo = result[i]["MtlItemNoDetail"].ToString();
                            var InventoryNo = result[i]["InventoryNoDetail"].ToString();
                            Double InventoryQty = 0;
                            dataRequest = productionDA.GetInventoryQty(MtlItemNo, InventoryNo);
                            jsonResponse = BaseHelper.DAResponse(dataRequest);
                            if (jsonResponse["status"].ToString() == "success")
                            {
                                var result1 = JObject.Parse(dataRequest)["data"];
                                if (result1.Count() > 0)
                                {
                                    InventoryQty = Convert.ToDouble(result1[0]["InventoryQty"].ToString());
                                }
                            }
                            #endregion

                            htmlDetail += @"<tr>
                                                <td style='width:4%;height:30px;text-align:center;'>[i]</td>
                                                <td style='width:23%;'>[MtlItemNoDetail]
                                                </td>
                                                <td style='width:23%;'>[MtlItemNameDetail]</td>
                                                <td style='width:5%;text-align:center;'>[InventoryNoDetail]</td>
                                                <td style='width:10%;text-align:center;'>[InventoryNameDetail]</td>
                                                <td style='width:5%;text-align:center;'>[SubstituteStatus]</td>
                                                <td style='width:10%;text-align:center;'>[DemandRequisitionQty]</td>
                                                <td style='width:13%;text-align:center;'>[InventoryQty]</td>
                                                <td style='width:5%;'></td>
                                            </tr>";
                            Double DemandRequisitionQty = Convert.ToDouble( result[i]["DemandRequisitionQty"].ToString());
                            htmlDetail = htmlDetail.Replace("[i]", (i + 1).ToString());
                            htmlDetail = htmlDetail.Replace("[MtlItemNoDetail]", result[i]["MtlItemNoDetail"].ToString());
                            htmlDetail = htmlDetail.Replace("[MtlItemNameDetail]", result[i]["MtlItemNameDetail"].ToString());
                            htmlDetail = htmlDetail.Replace("[InventoryNoDetail]", result[i]["InventoryNoDetail"].ToString());
                            htmlDetail = htmlDetail.Replace("[InventoryNameDetail]", result[i]["InventoryNameDetail"].ToString());
                            htmlDetail = htmlDetail.Replace("[SubstituteStatus]", result[i]["SubstituteStatus"].ToString() =="Y" ? "是":"否");
                            htmlDetail = htmlDetail.Replace("[DemandRequisitionQty]", DemandRequisitionQty.ToString("N"));
                            htmlDetail = htmlDetail.Replace("[InventoryQty]", InventoryQty.ToString("N"));
                        }
                        htmlText = htmlText.Replace("[Detail]", htmlDetail);
                    }
                    else
                    {
                        string htmlDetail = @"<tr>
                                                <td style='width:30px;height:30px;'></td>
                                                <td style='width:80px;'></td>
                                                <td style='width:160px;'>以下空白 / /</td>
                                                <td style='width:30px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:120px;'></td>
                                                <td style='width:50px;'></td>
                                                <td style='width:50px;'></td>
                                              </tr>";

                        for (int i = 0; i < pageNum - 1; i++)
                        {
                            htmlDetail += @"<tr>
                                                <td style='width:30px;height:30px;'></td>
                                                <td style='width:80px;'></td>
                                                <td style='width:160px;'>以下空白 / /</td>
                                                <td style='width:30px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:25px;'></td>
                                                <td style='width:120px;'></td>
                                                <td style='width:50px;'></td>
                                                <td style='width:50px;'></td>
                                            </tr>";
                        }

                        htmlText = htmlText.Replace("[Detail]", htmlDetail);
                    }
                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
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
                FileInfo file = new FileInfo(filePath);
                file.Delete();

                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = "MES2.0領料卡" + MoNo,
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

        #region //GetBarcodeLetteringPdf 刻字條碼列印
        public void GetBarcodeLetteringPdf(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request
                dataRequest = productionDA.GetBarcodeLetteringPdf(MoId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string htmlText = "";
                string MoNo = "";


                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    #region //html

                    if (result.Count() <= 0) throw new SystemException("此製令無法進行列印作業!");

                    var num = 1;
                    string htmlDetail = "";
                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/BarcodeLettering.html"));
                    MoNo = result[0]["MoNo"].ToString();
                    htmlText = htmlText.Replace("[MoNo]", result[0]["MoNo"].ToString());

                    var rowNum = 6;
                    string frameWidth = "1300px";
                    string phptpWidth = "100px";
                    string height = "70px";
                    var resultCont = result.Count();
                    var mod = resultCont % rowNum;
                    List<string> barcodeList = new List<string>();

                    foreach (var item in data)
                    {
                        var writer = new BarcodeWriter  //dll裡面可以看到屬性
                        {
                            Format = BarcodeFormat.QR_CODE,
                            Options = new QrCodeEncodingOptions //設定大小
                            {
                                Height = 10,
                                Width = 10,
                                PureBarcode = true,
                                Margin = 30
                            }

                        };
                        //產生QRcode
                        var img = writer.Write(Convert.ToString(item.MoItemPartId.ToString()));
                        string FileName = "BarcodeNo" + num;
                        Bitmap myBitmap = new Bitmap(img);
                        string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/LetteringBarcode/{0}.bmp", FileName));

                        myBitmap.Save(filePath, ImageFormat.Bmp);
                        barcodeList.Add(Convert.ToString(item.MoItemPartNo.ToString()));


                        if (num % rowNum == 0)
                        {
                            htmlDetail += @"<tr>";
                            for (var i = 1; i <= rowNum; i++)
                            {
                                htmlDetail += @"
                                                        < td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                            <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                                [BarcodeNo]
                                                        </td>
                                                        ";
                            }
                            htmlDetail += @"</tr>";
                            for (var i = num - rowNum + 1; i <= num; i++)
                            {
                                StringBuilder sb = new StringBuilder(htmlDetail);
                                int index = htmlDetail.IndexOf("[BarcodeNo]");
                                htmlDetail = htmlDetail.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                            }
                            for (var i= num - rowNum+1; i<= num; i++)
                            {
                                StringBuilder sb = new StringBuilder(htmlDetail);
                                int index = htmlDetail.IndexOf("[BarcodeNoImg]");
                                htmlDetail = htmlDetail.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/LetteringBarcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                            }
                            barcodeList.Clear();
                        }
                        if (mod > 0)
                        {
                            if (num == resultCont)
                            {
                                htmlDetail += @"<tr>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo1]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo2]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo3]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo4]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo5]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black;'>
                                                </td>
                                            </tr>
                                                ";
                                htmlDetail = htmlDetail.Replace("[BarcodeNo1]", barcodeList[0]);
                                htmlDetail = htmlDetail.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                                htmlDetail = htmlDetail.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                                htmlDetail = htmlDetail.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                                htmlDetail = htmlDetail.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                                for (var i = num - (num % rowNum) + 1; i <= num; i++)
                                {
                                    StringBuilder sb = new StringBuilder(htmlDetail);
                                    int index = htmlDetail.IndexOf("[BarcodeNoImg]");
                                    htmlDetail = htmlDetail.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/LetteringBarcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                                }
                                htmlDetail = htmlDetail.Replace("[BarcodeNoImg]", "");

                                barcodeList.Clear();
                            }
                        }


                        num++;
                    }
                    htmlText = htmlText.Replace("[BarcodeItem]", htmlDetail);
                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
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
                //FileInfo file = new FileInfo(filePath);
                //file.Delete();
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/LetteringBarcode"));

                FileInfo[] files = di.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName =" 【MES2.0】刻字條碼資料" + MoNo,
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

        #region //GetMoBoxPdf 模仁盒資訊列印(MES)
        public void GetMoBoxPdf(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - 流程卡資料
                dataRequest = productionDA.GetManufactureOrder(MoId, -1, "", "", "", "", -1, "", -1, "", "", "", "", -1, "", "", -1, -1);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string htmlText = "";
                string MoBoxNo = "";

                dynamic[] result = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                foreach (var item in result)
                {
                    var writer = new BarcodeWriter  //dll裡面可以看到屬性
                    {
                        Format = BarcodeFormat.CODE_128,
                        Options = new EncodingOptions //設定大小
                        { 
                            Height = 20,
                            Width = 150,
                            PureBarcode = true, //可設置是否顯示條碼底部內容
                            Margin = 0
                        }
                    };
                    //產生QRcode
                    var img = writer.Write(Convert.ToString(item.MoId));
                    string FileName = Convert.ToString(item.MoId);
                    Bitmap myBitmap = new Bitmap(img);
                    string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/MoBoxBarcode/{0}.bmp", FileName));
                    myBitmap.Save(filePath, ImageFormat.Bmp);


                    if (jsonResponse["status"].ToString() == "success")
                    {
                        MoBoxNo = result[0]["WoErpPrefix"].ToString() + "-" + result[0]["WoErpNo"].ToString() + "(" + result[0]["WoSeq"].ToString() + ")";

                        #region //html

                        if (result.Count() <= 0) throw new SystemException("此製令無法進行列印作業!");

                        string a = result[0]["MtlItemName"].ToString();

                        htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/MoBox.html"));
                        htmlText = htmlText.Replace("[CompanyName]", result[0]["CompanyName"].ToString());
                        htmlText = htmlText.Replace("[WoErpPrefix]", result[0]["WoErpPrefix"].ToString());
                        htmlText = htmlText.Replace("[WoErpNo]", result[0]["WoErpNo"].ToString());
                        htmlText = htmlText.Replace("[WoSeq]", result[0]["WoSeq"].ToString());
                        htmlText = htmlText.Replace("[MtlItemNo]", HttpUtility.HtmlEncode(result[0]["MtlItemNo"].ToString()));
                        //string mtlItemName = HttpUtility.HtmlEncode(result[0]["MtlItemName"].ToString());
                        string mtlItemName = result[0]["MtlItemName"].ToString();
                        int mtlItemLength = mtlItemName.Length;
                        if (mtlItemLength >= 25)
                        {
                            string firstPart = HttpUtility.HtmlEncode(mtlItemName.Substring(0, 20));
                            string secondPart = HttpUtility.HtmlEncode(mtlItemName.Substring(20));
                            htmlText = htmlText.Replace("[MtlItemName1]", "<font size='5'><b>" + firstPart + "</b></font>");
                            htmlText = htmlText.Replace("[MtlItemName2]", "<font size='5'><b>" + secondPart + "</b></font>");
                        }
                        else
                        {
                            htmlText = htmlText.Replace("[MtlItemName1]", "<font size='5'><b>" + HttpUtility.HtmlEncode(mtlItemName) + "</b></font>");
                            htmlText = htmlText.Replace("[MtlItemName2]", "");
                        }
                        htmlText = htmlText.Replace("[Quantity]", result[0]["Quantity"].ToString());
                        if (result[0]["CompanyName"].ToString() != "中揚光電")
                        {
                            if (result[0]["CustomerShortName"].ToString() =="")
                            {
                                htmlText = htmlText.Replace("[CustName]", "<pre>                </pre><b>【客戶】</b>");
                            }
                            else
                            {
                                htmlText = htmlText.Replace("[CustName]", "<pre>                </pre><b>【客戶】" + result[0]["CustomerShortName"].ToString()+ "</b>");
                            }
                        }
                        else {
                            htmlText = htmlText.Replace("[CustName]", "");
                        } 
                        htmlText = htmlText.Replace("[MoId]", Server.MapPath("~/PdfTemplate/MES/MoBoxBarcode/" + FileName + ".bmp").ToString());
                        htmlText = htmlText.Replace("[ExpectedEnd]", result[0]["ExpectedEnd"].ToString());
                        htmlText = htmlText.Replace("[WoRemark]", result[0]["Remark"].ToString() == null ? "" : "<font size='5'><b>" + result[0]["Remark"].ToString() + "</b></font>");

                        string htmlTemplate = htmlText;
                        #endregion
                    }

                }
                

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        // 設定邊距：左、右、上、下
                        float leftMargin = 20f;  // 左邊距
                        float rightMargin = 20f; // 右邊距
                        float topMargin = 10f;    // 上邊距
                        float bottomMargin = 10f; // 下邊距

                        using (Document document = new Document(PageSize.A6.Rotate(), leftMargin, rightMargin, topMargin, bottomMargin)) // 設定邊距
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

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = "MES2.0模仁盒" + MoBoxNo,
                    fileExtension = ".pdf"
                });
                #endregion

                #region//刪除模仁盒條碼
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/MoBoxBarcode"));
                FileInfo[] files = di.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
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

        #region //DownloadBarcodeLetteringPdf 下載刻字歷史紀錄PDF
        public void DownloadBarcodeLetteringPdf(string BarcodeNoList = "", string WoErpFullNo = "")
        {
            try
            {
                if (BarcodeNoList.Length <= 0) throw new SystemException("列印刻字列表不能為空!!");
                if (WoErpFullNo.Length <= 0) throw new SystemException("製令不能為空!!");

                WebApiLoginCheck("MfgOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request
                //dataRequest = productionDA.GetBarcodeLetteringPdf(MoItemPartIdList);
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string htmlText = "";
                #region //html

                string[] BarcodeNoList2 = BarcodeNoList.Split(',');

                if (BarcodeNoList2.Count() <= 0) throw new SystemException("請至少勾選一筆列印資料!!");

                var num = 1;
                string htmlDetail = "";
                htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/BarcodeLettering.html"));
                htmlText = htmlText.Replace("[MoNo]", WoErpFullNo);

                var rowNum = 6;
                string frameWidth = "1300px";
                string phptpWidth = "100px";
                string height = "70px";
                var resultCont = BarcodeNoList2.Count();
                var mod = resultCont % rowNum;
                List<string> barcodeList = new List<string>();

                foreach (var item in BarcodeNoList2)
                {
                    #region //取得刻字內容
                    dataRequest = productionDA.GetBarcodeAttribute02(item);
                    var result = JObject.Parse(dataRequest)["data"];

                    string Lettering = "";
                    foreach (var item2 in result)
                    {
                        Lettering = item2["ItemValue"].ToString();
                    }
                    #endregion

                    var writer = new BarcodeWriter  //dll裡面可以看到屬性
                    {
                        Format = BarcodeFormat.QR_CODE,
                        Options = new QrCodeEncodingOptions //設定大小
                        {
                            Height = 10,
                            Width = 10,
                            PureBarcode = true,
                            Margin = 30
                        }

                    };
                    //產生QRcode
                    var img = writer.Write(Convert.ToString(item.ToString()));
                    string FileName = "BarcodeNo" + num;
                    Bitmap myBitmap = new Bitmap(img);
                    string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/LetteringBarcode/{0}.bmp", FileName));

                    myBitmap.Save(filePath, ImageFormat.Bmp);
                    barcodeList.Add(Convert.ToString(Lettering.ToString()));


                    if (num % rowNum == 0)
                    {
                        htmlDetail += @"<tr>";
                        for (var i = 1; i <= rowNum; i++)
                        {
                            htmlDetail += @"
                                                        < td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                            <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                                [BarcodeNo]
                                                        </td>
                                                        ";
                        }
                        htmlDetail += @"</tr>";
                        for (var i = num - rowNum + 1; i <= num; i++)
                        {
                            StringBuilder sb = new StringBuilder(htmlDetail);
                            int index = htmlDetail.IndexOf("[BarcodeNo]");
                            htmlDetail = htmlDetail.Substring(0, index) + (barcodeList[(i - 1) % rowNum]) + htmlDetail.Substring(index + 11);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                        }
                        for (var i = num - rowNum + 1; i <= num; i++)
                        {
                            StringBuilder sb = new StringBuilder(htmlDetail);
                            int index = htmlDetail.IndexOf("[BarcodeNoImg]");
                            htmlDetail = htmlDetail.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/LetteringBarcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                        }
                        barcodeList.Clear();
                    }
                    if (mod > 0)
                    {
                        if (num == resultCont)
                        {
                            htmlDetail += @"<tr>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo1]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo2]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo3]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo4]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black; font-size:12px; text-align:center;'>
                                                    <img style='width: " + phptpWidth + @";' src='[BarcodeNoImg]' alt=''/>
                                                        [BarcodeNo5]
                                                </td>
                                                <td style='width: " + frameWidth + @";height: " + height + @"; padding:4px; border: 1px solid black;'>
                                                </td>
                                            </tr>
                                                ";
                            htmlDetail = htmlDetail.Replace("[BarcodeNo1]", barcodeList[0]);
                            htmlDetail = htmlDetail.Replace("[BarcodeNo2]", barcodeList.Count >= 2 ? barcodeList[1] : "");
                            htmlDetail = htmlDetail.Replace("[BarcodeNo3]", barcodeList.Count >= 3 ? barcodeList[2] : "");
                            htmlDetail = htmlDetail.Replace("[BarcodeNo4]", barcodeList.Count >= 4 ? barcodeList[3] : "");
                            htmlDetail = htmlDetail.Replace("[BarcodeNo5]", barcodeList.Count >= 5 ? barcodeList[4] : "");
                            for (var i = num - (num % rowNum) + 1; i <= num; i++)
                            {
                                StringBuilder sb = new StringBuilder(htmlDetail);
                                int index = htmlDetail.IndexOf("[BarcodeNoImg]");
                                htmlDetail = htmlDetail.Substring(0, index) + (Server.MapPath("~/PdfTemplate/MES/LetteringBarcode/BarcodeNo" + i + ".bmp").ToString()) + htmlDetail.Substring(index + 14);//指定替换的范围实现替换一次,并且指定范围中要只有一个替换的字符串
                            }
                            htmlDetail = htmlDetail.Replace("[BarcodeNoImg]", "");

                            barcodeList.Clear();
                        }
                    }


                    num++;
                }
                htmlText = htmlText.Replace("[BarcodeItem]", htmlDetail);
                #endregion

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
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
                //FileInfo file = new FileInfo(filePath);
                //file.Delete();
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/LetteringBarcode"));

                FileInfo[] files = di.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = " 【MES2.0】刻字條碼資料" + WoErpFullNo,
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

        #region //Print
        #region //PrintMoFullNo -- 製令標籤列印  -- Shintokuro 2024-11-04
        [HttpPost]
        public void PrintMoFullNo(string MoIdList = "", string PrintMachine = "", string ViewCompanyId = "", int PrintNum = -1)
        {
            try
            {
                WebApiLoginCheck("LotBarcoderPrintLabel", "print");

                if (MoIdList.Length <= 0) throw new SystemException("MES製令不可以為空,請重新確認!!");
                if (PrintNum <= 0) throw new SystemException("列印數量不可以為空,請重新確認!!");
                if (PrintMachine.Length <= 0) throw new SystemException("標籤機不可以為空,請重新確認!!");

                string numAllStr = "";

                int PrintCout = 0;
                string BarcodeNoStr = "";
                string LabelPath = "", line2;
                ILabel label;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤檔
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintMoFullNo\\Template\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                #endregion

                #region //標籤機
                label.PrintSettings.PrinterName = PrintMachine;
                #endregion

                #endregion

                #region //資料取得                
                dataRequest = productionDA.GetBarcodePrintData(MoIdList);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #region //資料操作
                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("製令設定不支援印製流程卡條碼標籤功能!!");
                    for (int n = 0; n < PrintNum; n++)
                    {
                        foreach (var item in result)
                        {
                            #region //製令資料取值
                            string StrMoId = item["MoId"].ToString();
                            string Quantity = item["Quantity"].ToString();
                            string MoNo = item["MoNo"].ToString();
                            string MtlItemNo = item["MtlItemNo"].ToString();
                            string MtlItemName = item["MtlItemName"].ToString();
                            string MtlItemSpec = item["MtlItemSpec"].ToString();
                            string BarcodeCtrl = item["BarcodeCtrl"].ToString();
                            string LotStatus = item["LotStatus"].ToString();
                            #endregion

                            #region //標籤檔賦值
                            label.Variables["MoNo"].SetValue(MoNo);
                            label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                            label.Variables["MtlItemName"].SetValue(MtlItemName);
                            label.Variables["MtlItemSpec"].SetValue(MtlItemSpec);
                            label.Variables["Quantity"].SetValue(Quantity);
                            #endregion

                            #region //列印
                            //等待两秒钟
                            Thread.Sleep(2000);
                            label.Print(1);
                            #endregion

                        }
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

        #region //PrintBarcode -- 流程卡條碼列印  -- Shintokuro 2024-08-06
        [HttpPost]
        public void PrintBarcode(string MoIdList = "", string PrintMachine = "", string ViewCompanyId = "", int PrintNum = -1)
        {
            try
            {
                WebApiLoginCheck("LotBarcoderPrintLabel", "print");

                if (MoIdList.Length <= 0) throw new SystemException("MES製令不可以為空,請重新確認!!");
                if (PrintNum <= 0) throw new SystemException("列印數量不可以為空,請重新確認!!");
                if (PrintMachine.Length <= 0) throw new SystemException("標籤機不可以為空,請重新確認!!");

                string numAllStr = "";

                int PrintCout = 0;
                string BarcodeNoStr = "";
                string LabelPath = "", line2;
                ILabel label;

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤檔
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintBarcode\\Template\\LabelName.txt", Encoding.Default);
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                #endregion

                #region //標籤機
                label.PrintSettings.PrinterName = PrintMachine;
                #endregion

                #endregion

                #region //資料取得                
                dataRequest = productionDA.GetBarcodePrintData(MoIdList);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #region //資料操作
                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("製令設定不支援印製流程卡條碼標籤功能!!");
                    for (int n = 0; n < PrintNum; n++)
                    {
                        foreach (var item in result)
                        {
                            string StrMoId = item["MoId"].ToString();
                            int Quantity = Convert.ToInt32(item["Quantity"].ToString());
                            string MoNo = item["MoNo"].ToString();
                            string MtlItemNo = item["MtlItemNo"].ToString();
                            string MtlItemName = item["MtlItemName"].ToString();
                            string MtlItemSpec = item["MtlItemSpec"].ToString();
                            string BarcodeCtrl = item["BarcodeCtrl"].ToString();
                            string LotStatus = item["LotStatus"].ToString();

                            if (BarcodeCtrl == "Y" && LotStatus == "N")
                            {
                                for (var i = 1; i <= Quantity; i++)
                                {
                                    string numStr = "";

                                    #region //標籤程式賦直列印
                                    string StrNo = string.Format("{0:0000}", i);

                                    string BarcodeNo = StrMoId + StrNo;

                                    label.Variables["BarcodeNo"].SetValue(BarcodeNo);
                                    label.Variables["MoNo"].SetValue(MoNo);
                                    label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                                    label.Variables["MtlItemName"].SetValue(MtlItemName);
                                    label.Variables["MtlItemSpec"].SetValue(MtlItemSpec);

                                    //等待两秒钟
                                    Thread.Sleep(2000);
                                    label.Print(1);
                                    numStr += BarcodeNo + "循環" + n + "</br>";
                                    numAllStr += numStr;
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("製令設定不支援印製流程卡條碼標籤功能!!");
                            }
                        }
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

        #endregion

        #region//FPR EIP API
        #region //GetMoSettingEIP 取得MES製令設定
        [HttpPost]
        [Route("api/CR/GetMoSetting")]

        public void GetMoSettingEIP(int MoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetMoSettingEIP(MoId
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

        #endregion

        #region //For批量開立製令
        #region //ExpandStructure 品號展開下階BOM物料
        [HttpPost]
        public void ExpandStructure(string UploadJson = "", string BomExpandFlag = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.ExpandStructure(UploadJson, BomExpandFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //BatchAddManufactureOrder 批量開立製令
        [HttpPost]
        public void BatchAddManufactureOrder(string UploadJson = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "update");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.BatchAddManufactureOrder(UploadJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetErpInventoryQty 取得庫存數量
        [HttpPost]
        public void GetErpInventoryQty(string MtlItemNo = "", string InventoryNo = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetErpInventoryQty(MtlItemNo, InventoryNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSingleErpInventoryQty 取得庫存數量
        [HttpPost]
        public void GetSingleErpInventoryQty(string MtlItemNo = "", string InventoryNo = "")
        {
            try
            {
                WebApiLoginCheck("MfgOrderManagement", "read");

                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.GetSingleErpInventoryQty(MtlItemNo, InventoryNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //PC平台測試
        #region //ExpandStructureForPC 品號展開下階BOM物料
        [HttpPost]
        public void ExpandStructureForPC(string MtlItemNo, string CompanyNo)
        {
            try
            {
                #region //Request
                productionDA = new ProductionDA();
                dataRequest = productionDA.ExpandStructureForPC(MtlItemNo, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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