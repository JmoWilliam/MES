using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;
using System.Reflection;
using ZXing;
using ZXing.Common;
using System.Drawing;
using System.Drawing.Imaging;

namespace Business_Manager.Controllers
{
    public class DeliveryController : WebController
    {
        private DeliveryDA deliveryDA = new DeliveryDA();

        #region //View
        public ActionResult DeliverySchedule()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DeliveryHistory()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DeliveryFinalizeChange()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DeliveryPlanning()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetDeliverySchedule 取得出貨排程資料
        [HttpPost]
        public void GetDeliverySchedule(string DeliveryStatus = "", string SoIds = "", string SoErpFullNo = "", string CustomerMtlItemNo = "", int CustomerId = -1
            , string MtlItemNo = "", int SalesmenId = -1, string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliverySchedule(DeliveryStatus, SoIds, SoErpFullNo, CustomerMtlItemNo, CustomerId, MtlItemNo, SalesmenId, StartDate, EndDate
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

        #region //GetDeliveryWipDetail 取得出貨排程製令資料
        [HttpPost]
        public void GetDeliveryWipDetail(int SoDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryWipDetail(SoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDeliveryFinalize 取得出貨定版資料
        [HttpPost]
        public void GetDeliveryFinalize(string StartDate = "", string EndDate = "", string ShipmentCustomer = "", string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "delivery-finalize");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryFinalize(StartDate, EndDate, ShipmentCustomer, MtlItemNo, MtlItemName
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

        #region //GetDeliveryDateLog 取得交期歷史紀錄
        [HttpPost]
        public void GetDeliveryDateLog(int SoDetailId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "date-log");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryDateLog(SoDetailId
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

        #region //GetDeliveryHistory 取得出貨歷史紀錄
        [HttpPost]
        public void GetDeliveryHistory(string StartDate = "", string EndDate = "", string CustomerIds = "", string DcIds = ""
            , string SoErpFullNo = "", string ItemValue = "",string MtlItemNo = "",string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeliveryHistory", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryHistory(StartDate, EndDate, CustomerIds, DcIds
                    , SoErpFullNo, ItemValue, MtlItemNo,MtlItemName
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

        #region //GetWipLink 取得製令綁定資料
        [HttpPost]
        public void GetWipLink(int SoDetailId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "wip-link");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetWipLink(SoDetailId, SearchKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDeliveryLetteringInfo 取得出貨歷史查詢之刻號資料
        [HttpPost]
        public void GetDeliveryLetteringInfo(int DoDetailId = -1, string ItemType = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryHistory", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryLetteringInfo(DoDetailId, ItemType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDoDetail 取得出貨單身資料
        [HttpPost]
        public void GetDoDetail(string DeliveryStatus = "", string DoErpFullNo = "", string SoIds = "", string SoErpFullNo = "", string CustomerMtlItemNo = "", int CustomerId = -1
            , string MtlItemNo = "", int SalesmenId = -1, string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {

            try
            {
                WebApiLoginCheck("DeliveryPlanning", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDoDetail(DeliveryStatus, DoErpFullNo, SoIds, SoErpFullNo, CustomerMtlItemNo, CustomerId, MtlItemNo, SalesmenId, StartDate, EndDate
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

        #region //GetSoDetail 取得出貨單身資料
        [HttpPost]
        public void GetSoDetail(int SoDetailId = -1, int SoId = -1, int MtlItemId = -1, string SoErpFullNo = "", string CustomerMtlItemNo = "", string TransferStatus = "", string SearchKey = "", string MtlItemIdIsNull = "", int CompanyId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {

            try
            {
                WebApiLoginCheck("DeliveryPlanning", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetSoDetail(SoDetailId, SoId, MtlItemId, SoErpFullNo, CustomerMtlItemNo, TransferStatus, SearchKey, MtlItemIdIsNull, CompanyId
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

        #region //GetDeliveryOrderDateLog 取得出貨日歷史紀錄資料
        [HttpPost]
        public void GetDeliveryOrderDateLog(int DoDetailId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {

            try
            {
                WebApiLoginCheck("DeliveryPlanning", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryOrderDateLog(DoDetailId
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

        #region //GetSoDetailForDeliveryOrder 取得訂單單身資料(取得可定版資料用)
        [HttpPost]
        public void GetSoDetailForDeliveryOrder(string SoErpFullNo = "")
        {

            try
            {
                WebApiLoginCheck("DeliveryPlanning", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetSoDetailForDeliveryOrder(SoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPendingOrders 取得待處理訂單資料
        [HttpPost]
        public void GetPendingOrders(string SoErpFullNo = "", string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {

            try
            {
                WebApiLoginCheck("DeliveryPlanning", "read");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetPendingOrders(SoErpFullNo, SearchKey
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
        #endregion

        #region //Add
        #region //AddDeliveryDateLog 交期歷史紀錄新增
        [HttpPost]
        public void AddDeliveryDateLog(int DoDetailId = -1, int SoDetailId = -1, string PcPromiseDate = "", string PcPromiseTime = ""
            , int CauseType = -1, int DepartmentId = -1, int SupervisorId = -1, string CauseDescription = "")
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "date-log");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.AddDeliveryDateLog(DoDetailId, SoDetailId, PcPromiseDate, PcPromiseTime, CauseType, DepartmentId, SupervisorId, CauseDescription);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeliveryFinalize 出貨定版資料新增
        [HttpPost]
        public void AddDeliveryFinalize(string Delieverys = "")
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "delivery-finalize");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.AddDeliveryFinalize(Delieverys);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeliveryFinalizeOrder 出貨定版資料新增(另一版本)
        [HttpPost]
        public void AddDeliveryFinalizeOrder(string Delieverys = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryPlanning", "add");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.AddDeliveryFinalizeOrder(Delieverys);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeliveryOrderDateLog 交期歷史紀錄新增(另一版本)
        [HttpPost]
        public void AddDeliveryOrderDateLog(int DoDetailId = -1, int SoDetailId = -1, string PcPromiseDate = "", string PcPromiseTime = ""
            , int CauseType = -1, int DepartmentId = -1, int SupervisorId = -1, string CauseDescription = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryPlanning", "add");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.AddDeliveryOrderDateLog(DoDetailId, SoDetailId, PcPromiseDate, PcPromiseTime, CauseType, DepartmentId, SupervisorId, CauseDescription);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeliveryQtyLog 新增更改出貨數量紀錄
        [HttpPost]
        public void AddDeliveryQtyLog(int DoDetailId = -1, int DoQty = -1
            , int CauseType = -1, int DepartmentId = -1, int SupervisorId = -1, string CauseDescription = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryPlanning", "add");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.AddDeliveryQtyLog(DoDetailId, DoQty, CauseType, DepartmentId, SupervisorId, CauseDescription);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateWipLink 訂單製令綁定
        [HttpPost]
        public void UpdateWipLink(int SoDetailId = -1, string WipNo = "", string LinkStatus = "")
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "wip-link");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.UpdateWipLink(SoDetailId, WipNo, LinkStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLotDeliveryFinalize 出貨定版異動
        [HttpPost]
        public void UpdateLotDeliveryFinalize(string DoDetails = "", string PcPromiseDate = "", string PcPromiseTime = ""
            , int CauseType = -1, int DepartmentId = -1, int SupervisorId = -1, string CauseDescription = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryFinalizeChange", "batch-update");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.UpdateLotDeliveryFinalize(DoDetails, PcPromiseDate, PcPromiseTime, CauseType, DepartmentId, SupervisorId, CauseDescription);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteDeliveryFinalize 出貨定版資料刪除
        [HttpPost]
        public void DeleteDeliveryFinalize(int DoDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "delivery-finalize");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.DeleteDeliveryFinalize(DoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //MailAdvice
        #region //DailyDeliveryWipUnLinkDetailMailAdvice 每日出貨未綁定製令明細
        [HttpPost]
        [Route("api/ERP/DailyDeliveryWipUnLinkDetail")]
        public void DailyDeliveryWipUnLinkDetailMailAdvice(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DailyDeliveryWipUnLinkDetailMailAdvice");
                #endregion

                #region //Request
                dataRequest = deliveryDA.DailyDeliveryWipUnLinkDetailMailAdvice(Company);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DailyDeliveryMailAdvice 每日出貨明細
        [HttpPost]
        [Route("api/ERP/DailyDelivery")]
        public void DailyDeliveryMailAdvice(string Company, string SecretKey)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "DailyDeliveryMailAdvice");
                #endregion

                #region //Request
                dataRequest = deliveryDA.DailyDeliveryMailAdvice(Company);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Download
        #region //Excel
        #region //DeliveryHistoryExcelDownload 出貨歷史紀錄匯出Excel
        public void DeliveryHistoryExcelDownload(string StartDate = "", string EndDate = "", string CustomerIds = "", string DcIds = "", string MtlItemNo="", string MtlItemName = ""
            , string SoErpFullNo = "", string ItemValue = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryHistory", "excel");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryHistory(StartDate, EndDate, CustomerIds, DcIds, MtlItemNo, MtlItemName
                    , SoErpFullNo, ItemValue
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
                    string excelFileName = "出貨歷史Excel檔" + StartDate + "_" + EndDate;
                    string excelsheetName = "出貨歷史紀錄";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "狀態", "出貨時間", "出貨客戶", "品名", "預計數量", "出貨數量", "生管備註", "訂單狀況", "訂單", "訂單備註", "客戶採單", "出貨工程", "【正常品】刻號", "【贈品】刻號", "【備品】刻號" };
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
                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DoStatus.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DoDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.DcShortName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.RegularQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.PickRegularQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.PcDoDetailRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.OrderSituation.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.SoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.SoDetailRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.CustomerPurchaseOrder.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // 設定水平對齊為左對齊
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.DeliveryProcess.ToString(); // 出貨工程
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.RegularItemValue != null ? item.RegularItemValue.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.FreebieItemValue != null ? item.FreebieItemValue.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.SpareItemValue != null ? item.SpareItemValue.ToString() : ""; 
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = " ";
                        }
                        #endregion

                        #region //設定
                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
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

                        #region //設定刻號欄寬
                        //worksheet.Column(11).Width = 50;
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

        #region //DeliveryFinalizeExcelDownload 出貨定版匯出Excel
        public void DeliveryFinalizeExcelDownload(string StartDate = "", string EndDate = "", string ShipmentCustomer = "", string MtlItemNo = "", string MtlItemName = "")
        {
            try
            {
                WebApiLoginCheck("DeliverySchedule", "delivery-finalize");

                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDeliveryFinalize(StartDate, EndDate, ShipmentCustomer, MtlItemNo, MtlItemName
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
                    string excelFileName = "出貨定版Excel檔" + StartDate + "_" + EndDate;
                    string excelsheetName = "出貨定版";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "出貨單號", "出貨時間", "出貨客戶", "品名"
                        , "訂單數量", "訂單欠數", "預出數量", "實揀數量"
                        , "出貨工程", "訂單狀況", "訂單單號", "寄送方式", "備註", "客戶採單" };
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
                        string nextCustomer = "";
                        string lastCustomer = "";
                        int XLColorIndex = 0;
                        int xLBordersIndex = 0;
                        foreach (var item in data)
                        {
                            nextCustomer = item.CustomerShortName;

                            if (nextCustomer != lastCustomer && lastCustomer != "")
                            {
                                XLColorIndex++;
                                xLBordersIndex = 1;

                                if (XLColorIndex > 9) XLColorIndex = 0;
                                if (xLBordersIndex > 1) xLBordersIndex = 0;
                            }
                            else
                            {
                                xLBordersIndex = 0;
                            }

                            for (int i = 1; i < 14; i++)
                            {
                                worksheet.Cell(rowIndex + 1, i).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(rowIndex + 1, i).Style.Border.TopBorderColor = XLColor.Black;
                                worksheet.Cell(rowIndex, i).Style.Border.BottomBorder = xLBorderSizes[xLBordersIndex];
                                worksheet.Cell(rowIndex, i).Style.Border.BottomBorderColor = xLBorderColors[xLBordersIndex];
                                worksheet.Cell(rowIndex + 1, i).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(rowIndex + 1, i).Style.Border.LeftBorderColor = XLColor.Black;
                                worksheet.Cell(rowIndex + 1, i).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(rowIndex + 1, i).Style.Border.RightBorderColor = XLColor.Black;
                                worksheet.Cell(rowIndex + 1, i).Style.Fill.BackgroundColor = xLColors[XLColorIndex];
                            }

                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DoDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.CustomerShortName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.SoRegularQty.ToString("N0");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = (Convert.ToInt32(item.SoRegularQty) - Convert.ToInt32(item.ExistPickQty)).ToString("N0");

                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.RegularQty.ToString("N0");
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.PickRegularQty.ToString("N0");
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Style = defaultStyle;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.DeliveryProcessName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.OrderSituation.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.SoErpFullNo.ToString() + '-' + item.SoSequence;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.DeliveryMethod.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.PcDoDetailRemark.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.CustomerPurchaseOrder.ToString();

                            lastCustomer = item.CustomerShortName;
                        }

                        #endregion

                        #region //設定
                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
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

        #region //DailyDeliveryWipUnLinkDetailExcelDownload 每日出貨未綁定製令明細匯出Excel
        public void DailyDeliveryWipUnLinkDetailExcelDownload(string CompanyNo, string StartDate = "", string EndDate = "")
        {
            try
            {
                #region //Request
                deliveryDA = new DeliveryDA();
                dataRequest = deliveryDA.GetDailyDeliveryWipUnLinkDetail(CompanyNo, StartDate, EndDate);
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
                    string excelFileName = "每日出貨未綁定製令明細Excel檔" + StartDate + "_" + EndDate;
                    string excelsheetName = "每日出貨未綁定製令明細";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "出貨日期", "出貨客戶", "出貨數量", "訂單", "訂單數量", "品號", "品名", "規格", "生管", "是否綁定" };
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
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.DoDate.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DcShortName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.PickQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.SoErpFullNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.SoQty.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.MtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.MtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.MtlItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.DepartmentName.ToString() + item.PcUserNo.ToString() + item.PcUserName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = Convert.ToBoolean(item.WipLink) ? "" : "未綁定";
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
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

        #region //PDF
        #region //PrintPickingCartonPdf 箱號內含條碼PDF匯出
        public void PrintPickingCartonPdf(string CartonBarcode = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryHistory", "pdf");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request
                dataRequest = deliveryDA.GetPickingCartonBarcode(CartonBarcode);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string htmlText = "";

                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    #region //html

                    if (result.Count() <= 0) throw new SystemException("此箱號查無資料,無法進行列印作業!");
                    string htmlDetail = "";
                    string TableDetail = "";
                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/PickingCartonBarcode.html"));

                    var resultCont = result.Count();
                    int PageNum = 50;
                    string tableTemplate = @"<tr>
                                                <td style='text-align:center;width: 21%;height: 20px;'>製造lot no</td>
                                                <td style='text-align:center;width:  6%;height: 20px;'>lot數</td>
                                                <td style='text-align:center;width: 23%;height: 20px;'>Barcode</td>
                                                <td style='text-align:center;width: 21%;height: 20px;'>製造lot no</td>
                                                <td style='text-align:center;width:  6%;height: 20px;'>lot數</td>
                                                <td style='text-align:center;width: 23%;height: 20px;'>Barcode</td>
                                            </tr>
                                            [TableDetail]
                                            <tr style='page-break-after: always;'></tr>";
                    string rowTemplate = @"<tr>[rowHtm]</tr>";

                    string tableItem = "";
                    string rowHtm = "";
                    string TdItem = "";

                    bool OddEven = true;//Odd:true Even:false
                    int num = 1;
                    foreach (var item in data)
                    {
                        var writer = new BarcodeWriter  //dll裡面可以看到屬性
                        {
                            Format = BarcodeFormat.CODE_128,
                            Options = new EncodingOptions //設定大小
                            {
                                Height = 18,
                                Width = 0,
                                PureBarcode = true,
                                Margin = 10
                            }
                        };
                        //產生QRcode
                        var img = writer.Write(Convert.ToString(item.BarcodeNo.ToString()));
                        string FileName = item.BarcodeNo;
                        Bitmap myBitmap = new Bitmap(img);
                        string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/PickingCartonBarcode/{0}.bmp", FileName));

                        myBitmap.Save(filePath, ImageFormat.Bmp);

                        if (OddEven)
                        {
                            rowHtm = rowTemplate;
                            TdItem = @"<td style='width: 21%;height: 20px;text-align:center;'>" + item.BarcodeNo + @"</td>
                                       <td style='width:  6%;height: 20px;text-align:center;'>" + item.BarcodeQty + @"</td>
                                       <td style='width: 23%;height: 20px;'>
                                           <img src='[BarcodeNoImg]' alt=''/>
                                       </td>
                                        ";
                            TdItem = TdItem.Replace("[BarcodeNoImg]", filePath);
                            OddEven = false;
                            if(num == resultCont)
                            {
                                TdItem += @"<td></td>
                                        <td></td>
                                        <td></td>
                                        ";

                                rowHtm = rowHtm.Replace("[rowHtm]", TdItem);
                                TableDetail += rowHtm;
                                rowHtm = "";
                                TdItem = "";
                                OddEven = true;
                            }
                            //if (result.Count() > 0) throw new SystemException("此箱號查無資料,無法進行列印作業!");
                        }
                        else
                        {
                            TdItem += @"<td style='width: 21%;height: 20px;text-align:center;'>" + item.BarcodeNo + @"</td>
                                        <td style='width:  6%;height: 20px;text-align:center;'>" + item.BarcodeQty + @"</td>
                                        <td style='width: 23%;height: 20px;'>
                                           <img src='[BarcodeNoImg]' alt=''/>
                                        </td>
                                        ";
                            TdItem = TdItem.Replace("[BarcodeNoImg]", filePath);

                            rowHtm = rowHtm.Replace("[rowHtm]", TdItem);
                            TableDetail += rowHtm;
                            rowHtm = "";
                            TdItem = "";
                            OddEven = true;
                        }
                        if(num % PageNum == 0)
                        {
                            tableItem = tableTemplate;
                            tableItem = tableItem.Replace("[TableDetail]", TableDetail);
                            htmlDetail += tableItem;
                            TableDetail = "";
                        }
                        num++;
                    }
                    if(TableDetail != "")
                    {
                        tableItem = tableTemplate;
                        tableItem = tableItem.Replace("[TableDetail]", TableDetail);
                        htmlDetail += tableItem;
                    }
                    
                    htmlText = htmlText.Replace("[BarcodeItem]", htmlDetail);
                    htmlText = htmlText.Replace("[PdfName]", "出貨箱(" + CartonBarcode + ")內含條碼資料");
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
                
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/PickingCartonBarcode"));

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
                    fileName = "出貨箱(" + CartonBarcode+ ")內含條碼資料",
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