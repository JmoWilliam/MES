using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using MESDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Data.SqlClient;

using System.Security.Cryptography;
using PDMDA;

namespace Business_Manager.Controllers
{
    public class RoutingController : WebController
    {
        private RoutingDA routingDA = new RoutingDA();

        #region //View
        public ActionResult RoutingManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AiRoutingManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MoProcessChangeManagment()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetRouting 取得途程資料
        [HttpPost]
        public void GetRouting(int RoutingId = -1, int ModeId = -1, string RoutingType = "", string RoutingName = "", int RoutingItemId = -1, string MtlItemNo = "", string MtlItemName = ""
            , string Status = "", string StartDate = "", string EndDate = "", int UserId = -1,string ProcessIds=""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read,constrained-data");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRouting(RoutingId, ModeId, RoutingType, RoutingName, RoutingItemId, MtlItemNo, MtlItemName
                    , Status, StartDate, EndDate, UserId, ProcessIds
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

        #region //GetRoutingProcess 取得途程製程資料
        [HttpPost]
        public void GetRoutingProcess(int RoutingProcessId = -1, int RoutingId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read,constrained-data");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRoutingProcess(RoutingProcessId, RoutingId
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

        #region //GetRoutingItem 取得途程品號資料
        [HttpPost]
        public void GetRoutingItem(int RoutingItemId = -1, int RoutingId = -1, string MtlItemNo = "", string MtlItemName = ""
            , int ControlId = -1, string RoutingItemConfirm = "", int MtlItemId = -1, int MoId = -1, string RoutingName = "", int ModeId = -1, string Edition = "", string RoutingConfirm = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read,constrained-data");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRoutingItem(RoutingItemId, RoutingId, MtlItemNo, MtlItemName
                    , ControlId, RoutingItemConfirm, MtlItemId, MoId, RoutingName, ModeId, Edition, RoutingConfirm, Status
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

        #region //GetRdDesignControl 取得研發設計圖版本控制資料
        [HttpPost]
        public void GetRdDesignControl(int ControlId = -1, int DesignId = -1, string Edition = "", string StartDate = "", string EndDate = "", int MtlItemId = -1, string ReleasedStatus = "", string MtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "version,constrained-data");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRdDesignControl(ControlId, DesignId, Edition, StartDate, EndDate, MtlItemId, ReleasedStatus, MtlItemNo
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

        #region //GetRoutingItemAttribute 取得途程品號加工屬性
        [HttpPost]
        public void GetRoutingItemAttribute(int RoutingItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRoutingItemAttribute(RoutingItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRoutingItemProcess 取得途程品號流程卡資料
        [HttpPost]
        public void GetRoutingItemProcess(int RoutingItemId = -1, string OrderBy = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRoutingItemProcess(RoutingItemId, OrderBy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRoutingProcessItem 取得途程製程屬性檔
        [HttpPost]
        public void GetRoutingProcessItem(int RoutingProcessItemId = -1, int RoutingProcessId = -1, string ItemNo = "", string TypeName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRoutingProcessItem(RoutingProcessItemId, RoutingProcessId, ItemNo, TypeName
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

        #region //GetRipQcItem 取得製程量測項目
        [HttpPost]
        public void GetRipQcItem(int ItemProcessId = -1, int QcItemId = -1, string itemNo = "", string itemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRipQcItem(ItemProcessId, QcItemId, itemNo, itemName
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

        #region //GetQcItem 取得加工參數設定資料
        [HttpPost]
        public void GetQcItem(int QcItemId = -1, string ItemNo = "", string ItemName = "", int ProcessId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetQcItem(QcItemId, ItemNo, ItemName, ProcessId
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

        #region //GetMoProcessChange 取得製令途程變更單
        [HttpPost]
        public void GetMoProcessChange(int MpcId = -1, int MoId = -1, string StartDocDate = "", string EndDocDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetMoProcessChange(MpcId, MoId, StartDocDate, EndDocDate, Status
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

        #region //GetMoProcess 取得製令製程
        [HttpPost]
        public void GetMoProcess(int MoId = -1, string ProcessNo = "", string ProcessName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetMoProcess(MoId, ProcessNo, ProcessName
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
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetMpcRoutingProcess(MpcRoutingProcessId, MpcId
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
                dataRequest = routingDA.GetMpcBarcode(MpcId, MpcBarcodeId, BarcodeNo
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

        #region //GetModeProcess 取得生產模式製程資料
        [HttpPost]
        public void GetModeProcess(int ProcessId = -1, string ProcessNo = "", string ProcessName = "", string Status = "", int ProdMode = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "read");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetModeProcess(ProcessId, ProcessNo, ProcessName, Status, ProdMode
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

        #region //GetRoutingProcessItem 取得途程製程資料屬性
        [HttpPost]
        public void GetRoutingProcessItemList(int RoutingProcessId = -1, int RoutingId = -1, string ProcessIdList = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "read,constrained-data");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.GetRoutingProcessItemList(RoutingProcessId, RoutingId, ProcessIdList
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


        #region //GetRdDesign 取得研發設計圖資料 --- copy it from Drawing Controller by Xuan and MarkChen
        [HttpPost]
        public void GetRdDesign(int DesignId = -1, string CustomerMtlItemNo = "",string MtlItemNo="", int MtlItemId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "read,constrained-data");

                #region //Request
                var drawingDA = new DrawingDA();
                dataRequest = drawingDA.GetRdDesign(DesignId, CustomerMtlItemNo, MtlItemNo, MtlItemId
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
        #region //AddRouting 新增途程資料
        [HttpPost]
        public void AddRouting(string RoutingName = "", string RoutingType = "", int ModeId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddRouting(RoutingName, RoutingType, ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRoutingProcess 新增途程製程資料
        [HttpPost]
        public void AddRoutingProcess(int RoutingId = -1, string ProcessData = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddRoutingProcess(RoutingId, ProcessData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBatchRoutingProcess 批量新增途程製程資料
        [HttpPost]
        public void AddBatchRoutingProcess(int RoutingId = -1, string UploadJson = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddBatchRoutingProcess(RoutingId, UploadJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRoutingItem 新增途程品號資料
        [HttpPost]
        public void AddRoutingItem(int RoutingId = -1, int ControlId = -1, int MtlItemId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddRoutingItem(RoutingId, ControlId, MtlItemId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBatchRoutingItem 新增途程品號資料
        [HttpPost]
        public void AddBatchRoutingItem(int RoutingId = -1, string UploadJson = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddBatchRoutingItem(RoutingId, UploadJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddCopyRoutingItem 複製途程品號資料
        [HttpPost]
        public void AddCopyRoutingItem(int RoutingItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddCopyRoutingItem(RoutingItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddCopyRouting 複製途程資料
        [HttpPost]
        public void AddCopyRouting(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddCopyRouting(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRoutingProcessItem 新增途程製程屬性檔
        [HttpPost]
        public void AddRoutingProcessItem(int RoutingProcessId = -1, string ItemNo = "", string ChkUnique = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddRoutingProcessItem(RoutingProcessId, ItemNo, ChkUnique);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRipQcItem 新增途程製程屬性檔
        [HttpPost]
        public void AddRipQcItem(string RipQcItemData = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.AddRipQcItem(RipQcItemData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.AddMoProcessChange(MoId, MpcUserId, DocDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        //#region //AddMpcRoutingProcess 新增途程製程資料
        //[HttpPost]
        //public void AddMpcRoutingProcess(int MpcId = -1, string ProcessData = "")
        //{
        //    try
        //    {
        //        WebApiLoginCheck("MoProcessChangeManagment", "add");

        //        #region //Request
        //        dataRequest = routingDA.AddMpcRoutingProcess(MpcId, ProcessData);
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
        //#endregion

        #region //AddMpcRoutingProcessAll 新增途程製程資料
        [HttpPost]
        public void AddMpcRoutingProcessAll(int MpcId = -1, int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                dataRequest = routingDA.AddMpcRoutingProcessAll(MpcId, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.AddMpcBarcode(MpcId, BarcodeIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateRouting 更新途程資料
        [HttpPost]
        public void UpdateRouting(int RoutingId = -1, string RoutingName = "", string RoutingType = "", int ModeId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRouting(RoutingId, RoutingName, RoutingType, ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingStatus 更新途程啟用狀態
        [HttpPost]
        public void UpdateRoutingStatus(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "status-switch");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingStatus(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingProcess 更新途程製程資料
        [HttpPost]
        public void UpdateRoutingProcess(int RoutingProcessId = -1, int RoutingId = -1, int ProcessId = -1, string ProcessAlias = ""
            , string DisplayStatus = "", string NecessityStatus = "", string ProcessCheckStatus = "", string ProcessCheckType = "", string PackageFlag = "", string ConsumeFlag = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingProcess(RoutingProcessId, RoutingId, ProcessId, ProcessAlias
                    , DisplayStatus, NecessityStatus, ProcessCheckStatus, ProcessCheckType, PackageFlag, ConsumeFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateNecessityStatus 更新必要過站狀態
        [HttpPost]
        public void UpdateNecessityStatus(int RoutingProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "necessity-status");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateNecessityStatus(RoutingProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingProcessSort 更新途程製程排序
        [HttpPost]
        public void UpdateRoutingProcessSort(string RoutingProcessSort = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "sort");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingProcessSort(RoutingProcessSort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingItem 更新途程品號資料
        [HttpPost]
        public void UpdateRoutingItem(int RoutingItemId = -1, int RoutingId = -1, int ControlId = -1, int MtlItemId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingItem(RoutingItemId, RoutingId, ControlId, MtlItemId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingItemStatus 更新途程品號狀態
        [HttpPost]
        public void UpdateRoutingItemStatus(int RoutingItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "status-switch");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingItemStatus(RoutingItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingItemAttribute 更新途程品號加工屬性資料
        [HttpPost]
        public void UpdateRoutingItemAttribute(string AttributeData = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingItemAttribute(AttributeData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ConfirmRouting 確認途程
        [HttpPost]
        public void ConfirmRouting(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingConfirm(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UnbindConfirmRouting 解除確認途程
        [HttpPost]
        public void UnbindConfirmRouting(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UnbindConfirmRouting(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingItemConfirm 確認途程品號
        [HttpPost]
        public void UpdateRoutingItemConfirm(int RoutingItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingItemConfirm(RoutingItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingItemProcess 更新加工流程卡內容
        [HttpPost]
        public void UpdateRoutingItemProcess(string RoutingItemProcessData = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingItemProcess(RoutingItemProcessData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDisplayStatus 更新是否顯示在流程卡狀態
        [HttpPost]
        public void UpdateDisplayStatus(int RoutingProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "status-switch");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateDisplayStatus(RoutingProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingProcessItem 更新途程製程屬性檔
        [HttpPost]
        public void UpdateRoutingProcessItem(int RoutingProcessItemId = -1, int RoutingProcessId = -1, string ItemNo = "", string ChkUnique = "")
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "update");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.UpdateRoutingProcessItem(RoutingProcessItemId, RoutingProcessId, ItemNo, ChkUnique);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.UpdateMpcIdStatus(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.UpdateMpcRoutingProcessSort(MpcId, RoutingProcessSort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteRoutingProcess -- 刪除途程製程資料
        [HttpPost]
        public void DeleteRoutingProcess(int RoutingProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "delete");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.DeleteRoutingProcess(RoutingProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRoutingItem -- 刪除途程品號資料
        [HttpPost]
        public void DeleteRoutingItem(int RoutingItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "delete");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.DeleteRoutingItem(RoutingItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRoutingProcessItem -- 刪除途程製程屬性資料
        [HttpPost]
        public void DeleteRoutingProcessItem(int RoutingProcessItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "delete");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.DeleteRoutingProcessItem(RoutingProcessItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRipQcItem -- 刪除途程製程加工屬性資料
        [HttpPost]
        public void DeleteRipQcItem(int ItemProcessId = -1, int QcItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "delete");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.DeleteRipQcItem(ItemProcessId, QcItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRouting -- 刪除途程資料
        [HttpPost]
        public void DeleteRouting(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "delete");

                #region //Request
                routingDA = new RoutingDA();
                dataRequest = routingDA.DeleteRouting(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.DeleteMoProcessChange(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.DeleteMpcBarcode(MpcBarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.DeleteMpcBarcodeAll(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMpcRoutingProcess 刪除製令途程變更單
        [HttpPost]
        public void DeleteMpcRoutingProcess(int MpcRoutingProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("MoProcessChangeManagment", "detail");

                #region //Request
                dataRequest = routingDA.DeleteMpcRoutingProcess(MpcRoutingProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                dataRequest = routingDA.DeleteMpcRoutingProcessAll(MpcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region // 呼叫自動流程卡AI
        [HttpPost]
        public string AiRunCard(string arg1, string arg2)
        {
            string result = "", error = "";
            try
            {
                //共夾路徑
                //FilePath = @"\\192.168.20.199\mes_data\CUST_RD\JMO\RdDesign\TOMANDJERRY-001\M0845CA-P3-R2-3993.dxf";

                string python_path = "";
                string python_name = "main_for_MES";
                switch (python_name)
                {
                    case "main_for_MES":
                        python_path = @"‪‪C:\Python\Python36\Data\AiRought\AiRought\main_for_MES\main_for_MES.py";
                        break;
                    case "GM_for_MES":
                        python_path = @"C:\Projects\python\GM_V2_MES\GM_AP_MES_v2.py";
                        break;
                    case "temp":
                        python_path = @"C:\Projects\python\v2.4_MES\temp.py";
                        break;
                    case "main_for_MES_v26":
                        python_path = @"C:\Projects\python\v2.6_MES\main_for_MES.py";
                        break;
                    default:
                        throw new Exception(python_name + " not found");
                }

                ProcessStartInfo info = new ProcessStartInfo
                {
                    Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\"",python_path,arg1,arg2),
                    CreateNoWindow = true,
                    FileName = @"C:\Python\Python36\python.exe",    
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                };

                using (Process process = Process.Start(info))
                {
                    using (StreamReader sr_output = process.StandardOutput)
                    {
                        result = sr_output.ReadToEnd().Replace("\r", "").Replace("\n", "");
                    }
                }
                return result;
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            return result;
        }

        public async Task<string> ApiPostAutoRunCard(string file_id)
        {
            string res;
            #region             
            var clientUrl = "http://192.168.20.97/WorkManagement/Home/api_retrieve_file";

            using (var client = new HttpClient())
            {
                using (var content = new MultipartFormDataContent())
                {
                    content.Add(new StringContent(file_id), "\"file_id\"");
                    using (var message = await client.PostAsync(clientUrl, content))
                    {
                        //取回結果
                        var input = await message.Content.ReadAsStringAsync();
                        JObject obj = (JObject)JsonConvert.DeserializeObject(input);


                        #region //Response
                        JObject jRes = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "success",
                            date = input
                        });
                        res = jRes.ToString();
                        #endregion
                    }
                };
            }
            #endregion
            return res;
        }
        #endregion

        #region // 呼叫 人工智慧研發部 API 自動新建 一組全套或半套途程 by MarkChen 2023.08.21
        [HttpPost]
        public async Task<string> ApiAiAutoRouting(int MtlItemId, int ControlId)
        {
            #region //[開始] NOTE by MarkChen, 使用 工廠模式 (Factory Pattern), 將複雜部份推到 DA
            var flow = new RoutingDA.ResultFactory().Create(MtlItemId, ControlId);
            #endregion

            #region //Step1 確認 研發設計圖版本控制 基礎資料, 是否適合調用 人工智慧研發部 API
            //CASE_A_01 : 查無 RdDesginControl
            if (flow.Step1 == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_01_RdDesignControl_NOT_FOUND;
                return JsonConvert.SerializeObject(flow); ;
            }

            //"CASE_A_02 : 無 FileId",
            if (flow.Step1.Cad2DFile == null)
            {
                flow.Status = RoutingDA.Cases.CASE_A_02_NO_FILEID;
                return JsonConvert.SerializeObject(flow); ;
            }
            flow.Input.FileId =(int)flow.Step1.Cad2DFile;
            #endregion

            #region //Step2 調用 人工智慧研發部 API, 待命調整為內部調用 或 以實際路徑調用
            using (HttpClient httpClient = new HttpClient())
            {
                var loginUri = "http://192.168.20.97/WorkManagement/Home/tx_login_check";

                var loginContent = new MultipartFormDataContent();
                loginContent.Add(new StringContent("{\"job_no\":\"mesapi\",\"password\":\"JmoAdmin123456\",\"remember_me\":true}"), "DATA");

                var loginResponse = await httpClient.PostAsync(loginUri, loginContent);
                if (!loginResponse.IsSuccessStatusCode)
                {
                    flow.Status = RoutingDA.Cases.CASE_B_01_NOT_ABLE_TO_VISIT_AI_API;
                    return JsonConvert.SerializeObject(flow); ;
                }
                string content = await loginResponse.Content.ReadAsStringAsync();
                JObject jsonResponse = JObject.Parse(content);

                string status = (string)jsonResponse["status"];
                string msg = (string)jsonResponse["msg"];
                string retdata = (string)jsonResponse["retdata"];

                if (status != "success")
                {
                    flow.Status = RoutingDA.Cases.CASE_B_02_NOT_ABLE_TO_LOGIN_AI_API;
                    flow.Debug = msg;
                    return JsonConvert.SerializeObject(flow); ;
                }

                var requestUri = "http://192.168.20.97/WorkManagement/Home/api_retrieve_file";

                var formContent = new MultipartFormDataContent();
                formContent.Add(new StringContent(flow.Step1.Cad2DFile.ToString()), "file_id");

                var response = await httpClient.PostAsync(requestUri, formContent);
                if (!response.IsSuccessStatusCode)
                {
                    flow.Status = RoutingDA.Cases.CASE_B_03_FAILED_TO_GET_PARSE;
                    return JsonConvert.SerializeObject(flow); ;
                }

                // NOTE by Mark, 合併簡潔寫法, 直接將 人工智慧研發部 解析的 JSON 存入 Step2
                flow.Step2 = JsonConvert.DeserializeObject<RoutingDA.Root>(await response.Content.ReadAsStringAsync());

                if (flow.Step2.Python_Ret.Status != "success")
                {
                    flow.Status = RoutingDA.Cases.CASE_B_04_PARSED_RESULT_NOT_USEFUL;
                    flow.Debug = "有解析, 但不能為本案所用, 目前只限 mode = 1,  JMO-A-001	模仁加工";
                    return JsonConvert.SerializeObject(flow); ;
                }
            }
            #endregion

            #region //Step3 判斷新增模式是全套途程或半套途程
            // NOTE by Mark, 合併簡潔寫法, 此段仍可優化
            flow.Step3.RoutingName = routingDA.GetRoutingNameV3(flow.Step2, (int)flow.Step1.Cad2DFile);
            flow.Step3.SameNameRouting = routingDA.GetRoutingByName(flow.Step3.RoutingName);

            //判斷是否已有 同名製程, 
            //  發生情況, 例如 /ApiAiAutoRouting?MtlItemId=264065&ControlId=3987 , Call了兩次
            if (flow.Step3.SameNameRouting.RoutingId > 0)
            {
                flow.Status = RoutingDA.Cases.CASE_C_01_EXISTING_ROUTINGNAME; //已有同名製程
                flow.Step3.IsAnySameRoutingName = true;
                return JsonConvert.SerializeObject(flow); ;
            }

            //在JSON, 可以看起來 乾淨! 不會展開沒有值的 { }
            flow.Step3.SameNameRouting = null;
            flow.Step3.IsAnySameRoutingName = false;

            // DOING... 使用密碼學比較有沒有相同途程
            RoutingDA.IsExistingRoutingResponse r1 = routingDA.IsExistingRoutingAdvAsync(flow.Step2, (int)flow.Step1.Cad2DFile);
            flow.Step3.EntireCompanyRoutingCnt = r1.EntireCompanyRoutingCnt;

            flow.Step3.ProcessCnt = r1.ProcessCnt;
            flow.Step3.SameProcessCntRoutingCnt = r1.SameProcessCntRoutingCnt;
            flow.Step3.strGoingToAdd = r1.strGoingToAdd;
            flow.Step3.IsAnySameProcessList = r1.IsAnySameProcessList;
            flow.Step3.SameProcessListDifferentRoutingName = r1.SameProcessListDifferentRoutingName;
            #endregion

            #region //Step4 執行 新增 全套途程 或 半套途程
            if (!flow.Step3.IsAnySameProcessList)
            {
                //新增  上邊 右邊  左邊 
                flow.Status = RoutingDA.Cases.CASE_D_01_TO_CREATE_ENTIRE_ROUTING;
                flow.Debug = "新增  上邊 右邊  左邊 ";

                flow.Step4.DebugMsg = "新增  上邊 右邊  左邊, calling  ExecuteTransaction_CASE_D_01... ";
                flow = routingDA.ExecuteTransaction_CASE_D_01(flow, MtlItemId, ControlId);
            }
            else
            {
                //只新增 左邊 
                flow.Status = RoutingDA.Cases.CASE_D_02_TO_CREATE_ROUTING_ITEM_ONLY;
                flow.Debug = "只新增 左邊 .";

                flow.Step4.DebugMsg = "新增    左邊, calling  ExecuteTransaction_CASE_D_02... ";
                flow = routingDA.ExecuteTransaction_CASE_D_02(flow, MtlItemId, ControlId);
            }
            #endregion

            #region //[結束] 回傳結果, 可以 不顯示DEBUG 細節
            //Clean Up Debug info
            flow.Cases = null;
            flow.Step1 = null;
            flow.Step2 = null;
            flow.Step3 = null;
            flow.Step4 = null;

            //flow.Step4.RoutingProcessList = null;
            return JsonConvert.SerializeObject(flow); ;
            #endregion
        }
        #endregion

        #region // 呼叫 人工智慧研發部 API 自動新建 一組全套或半套途程。按秉丞指導:不使用 async, 業務邏輯全部改到DA,  by MarkChen 2023.08.24
        [HttpPost]
        public string ApiAiAutoRoutingV2(int MtlItemId, int ControlId)
        {
            var res = routingDA.ApiAiAutoRoutingV2(MtlItemId, ControlId);
            return res;

        }
        #endregion
        #region // 呼叫 人工智慧研發部 API 自動新建 一組全套或半套途程。按經理 08/25 指導: 不可直接按 JSON 自動生成完畢,要多一個動作讓用戶可以做[途程製程維護], 這 Step1, 只做到取了可用的JSON讓用戶人工維護  by MarkChen 2023.08.25
        [HttpPost]
        public string ApiAiAutoRoutingV2Step1(int MtlItemId, int ControlId)
        {
            var res = routingDA.ApiAiAutoRoutingV2Step1(MtlItemId, ControlId);
            return res;
        }
        #endregion

        #region // 呼叫 人工智慧研發部 API 自動新建 一組全套或半套途程。Step2! 按照 09/01 部門會議經理指示, 以目前UI和功能,[確定]後直接生成整套途程即可。不再處理製程順序或新增製程後的刪掉等功能 . NOTE by Mark 09/01
        [HttpPost]
        public string ApiAiAutoRoutingV2Step2(int MtlItemId, int ControlId, string ProcessJsonString)
        {
            var res = routingDA.ApiAiAutoRoutingV2Step2(MtlItemId, ControlId, ProcessJsonString);
            return res;
        }

        #endregion

        #endregion

        #region//Download
        #region //Pdf
        #region //GetFlowCardPdf 流程卡預覽
        public void GetPreFlowCardPdf(int RoutingItemId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - 流程卡資料
                dataRequest = routingDA.GetPreFlowCardPdf(RoutingItemId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string MoNo = "";
                string MtlItemName = "";
                string LotStatus = "";
                string BarcodeCtrl = "";
                string htmlText = "";
                string CompanyName = "";
                int LotQty = 0;

                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("此製令無法進行列印作業!可能因圖層未設定");
                    #region //html
                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/Routing/FlowCard.html"));

                    MtlItemName = result[0]["MtlItemName"].ToString();
                    htmlText = htmlText.Replace("[CompanyName]", "流程卡預覽");
                    htmlText = htmlText.Replace("[MoNo]", "");
                    htmlText = htmlText.Replace("[MtlItemNo]", result[0]["MtlItemNo"].ToString());
                    htmlText = htmlText.Replace("[MtlItemName]", result[0]["MtlItemName"].ToString());
                    htmlText = htmlText.Replace("[MtlItemSpec]", result[0]["MtlItemSpec"].ToString());
                    htmlText = htmlText.Replace("[InventoryName]", "");
                    htmlText = htmlText.Replace("[UserName]", "");
                    htmlText = htmlText.Replace("[Quantity]", "");
                    htmlText = htmlText.Replace("[ExpectedEnd]", "");
                    htmlText = htmlText.Replace("[PlanQty]", "");
                    htmlText = htmlText.Replace("[Edition]", result[0]["Edition"].ToString());
                    htmlText = htmlText.Replace("[CustomerDwgNo]", result[0]["CustomerMtlItemNo"].ToString());
                    htmlText = htmlText.Replace("[BarcodeCtrl]", "");
                    htmlText = htmlText.Replace("[OQcCheckType]", "");

                    string htmlTemplate = htmlText;
                    var pageNum = 12;
                    int totalRecords = result.Count();
                    int totalPages = totalRecords / pageNum + (totalRecords % pageNum != 0 ? 1 : 0);

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

                            string Remark = result[i]["Remark"].ToString().Replace("<", " &#60; ");
                            if (Remark == "")
                            {
                                Remark = "";
                            }
                            else {
                                Remark= Remark.Replace("\n", "<br/>");
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

                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = "【MES2.0】流程卡" + MtlItemName,
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
        #endregion
    }
}