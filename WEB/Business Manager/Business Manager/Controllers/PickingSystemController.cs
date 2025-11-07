using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class PickingSystemController : WebController
    {
        private PickingSystemDA pickingSystemDA = new PickingSystemDA();

        #region //View
        #region//Sample
        public ActionResult IndexSample()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ShipList()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ShipListSI()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ShipListSI02()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ShipListSE()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ShipListSE02()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ListPicknot()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ListPicking()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ListShipping()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ListShipped()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult PickTask01()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult PickTask02()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult PickTask03()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult PickTask04unedit()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ChangeBox()
        {
            ViewLoginCheck();

            return View();
        }
       
        public ActionResult ChangeBoxEdit()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxList()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxList02()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxTask01()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxTask02()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxTask03()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxTask04()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult BoxAdd()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        public ActionResult Login()
        {
            return View();
        }

        public ActionResult Index()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult CartonList()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult CartonForm()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult DeliveryList()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult DeliveryDetail()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult PickingItem()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult CartonItem()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult ChangeCarton()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult PickingTask()
        {
            PickingLoginCheck();

            return View();
        }

        public ActionResult PackagingBarcode()
        {
            PickingLoginCheck();

            return View();
        }

        #endregion

        #region //Get
        #region //GetCartonBarcode 取得物流箱條碼
        [HttpPost]
        public void GetCartonBarcode()
        {
            try
            {
                PickingApiLoginCheck();

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    result = BaseHelper.RandomCode(6)
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

        #region //GetCartonDetail 取得所有物流箱資料
        [HttpPost]
        public void GetCartonDetail(int DoId = -1, int DoDetailId = -1, int PickingCartonId = -1, string CartonBarcode = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetCartonDetail(DoId, DoDetailId, PickingCartonId, CartonBarcode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPickingCarton 取得物流箱資料(不含有出貨單)
        [HttpPost]
        public void GetPickingCarton(int PickingCartonId = -1, int PackingId = -1, string CartonName = "", string CartonBarcode = "", string CartonStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetPickingCarton(PickingCartonId, PackingId, CartonName, CartonBarcode, CartonStatus
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

        #region //GetChangeCarton 取得換箱資料(含有出貨單)
        [HttpPost]
        public void GetChangeCarton(string CartonBarcode = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetChangeCarton(CartonBarcode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPickingItem 取得揀貨物件資料
        [HttpPost]
        public void GetPickingItem(int DoId = -1, int PickingItemId = -1, int PickingCartonId = -1, int SoDetailId = -1, string PickingItemIds = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetPickingItem(DoId, PickingItemId, PickingCartonId, SoDetailId, PickingItemIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDeliveryOrder 取得出貨單資料
        [HttpPost]
        public void GetDeliveryOrder(int DoId = -1, string TypeTwo = "", string DoErpFullNo = "", string CartonBarcode = "", string Status = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetDeliveryOrder(DoId, TypeTwo, DoErpFullNo, CartonBarcode, Status, StartDate, EndDate
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

        #region //GetDoDetail 取得出貨單物件資料
        [HttpPost]
        public void GetDoDetail(int DoDetailId = -1, int DoId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetDoDetail(DoDetailId, DoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCartonItem 取得物流箱物件資料
        [HttpPost]
        public void GetCartonItem(int DoId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetCartonItem(DoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPickingTask 取得揀貨中物件資料
        [HttpPost]
        public void GetPickingTask(int DoDetailId = -1, int PickingCartonId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetPickingTask(DoDetailId, PickingCartonId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPackageInBarcode -- 取得包裝內容物條碼 
        [HttpPost]
        public void GetPackageInBarcode(string BarcodeNo = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetPackageInBarcode(BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion


        #region //GetDoDetailForPackage -- 取得出貨單物件資料(FOR PACKAGE)
        [HttpPost]
        public void GetDoDetailForPackage(int DoId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetDoDetailForPackage(DoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPickitem --  取得已揀貨條碼(FOR PACKAGE)
        [HttpPost]
        public void GetPickitem(int MtlItemId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetPickitem(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDeliveryLotList 取得揀貨中物件資料
        [HttpPost]
        public void GetDeliveryLotList(int DoDetailId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetDeliveryLotList(DoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPickingLot 取得已揀貨無條碼批號
        [HttpPost]
        public void GetPickingLot(int PickingItemId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.GetPickingLot(PickingItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddPickingCarton 物流箱資料新增
        [HttpPost]
        public void AddPickingCarton(int PackingId = -1, string CartonName = "", string CartonBarcode = "", string CartonRemark = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.AddPickingCarton(PackingId, CartonName, CartonBarcode, CartonRemark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPickingItem 物流箱資料新增
        [HttpPost]
        public void AddPickingItem(int DoDetailId = -1, int PickingCartonId = -1, string InputStatus = ""
            , string Barcode = "", int NoBarcodeQty = -1, int SubstituteQty = -1, string ItemType = ""
            , string NoBarcodeLotNumber = "", string SubstituteLotNumber = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.AddPickingItem(DoDetailId, PickingCartonId, InputStatus
                    , Barcode, NoBarcodeQty, SubstituteQty, ItemType
                    , NoBarcodeLotNumber, SubstituteLotNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPackagePickingItem -- 包裝條碼揀貨物件資料新增
        [HttpPost]
        public void AddPackagePickingItem(int DoDetailId = -1, int PickingCartonId = -1, string InputStatus = ""
            , string Barcode = "", int NoBarcodeQty = -1, int SubstituteQty = -1, string ItemType = "", string PickJson = ""
            )
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.AddPackagePickingItem(DoDetailId, PickingCartonId, InputStatus
                    , Barcode, NoBarcodeQty, SubstituteQty, ItemType, PickJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdatePickingCarton 物流箱資料更新
        [HttpPost]
        public void UpdatePickingCarton(int PickingCartonId = -1, int PackingId = -1, string CartonName = "", string CartonBarcode = "", string CartonRemark = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.UpdatePickingCarton(PickingCartonId, PackingId, CartonName, CartonBarcode, CartonRemark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePickingItem 揀貨物件資料更新(物件換箱)
        [HttpPost]
        public void UpdatePickingItem(string PickingItemIds = "", string CartonBarcode = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.UpdatePickingItem(PickingItemIds, CartonBarcode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDeleiveryStatus 出貨單狀態更新
        [HttpPost]
        public void UpdateDeleiveryStatus(string DoIds = "", string Status = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.UpdateDeleiveryStatus(DoIds, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCartonWeight 物流箱重量更新
        [HttpPost]
        public void UpdateCartonWeight(int PickingCartonId = -1, double TotalWeight = 0.0, int UomId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.UpdateCartonWeight(PickingCartonId, TotalWeight, UomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDoCarton 出貨單物流箱更新(綁定物流箱)
        [HttpPost]
        public void UpdateDoCarton(string CartonBarcode = "", int DoId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.UpdateDoCarton(CartonBarcode, DoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeletePickingCarton 物流箱資料刪除
        [HttpPost]
        public void DeletePickingCarton(string PickingCartonIds = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.DeletePickingCarton(PickingCartonIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePickingItem 揀貨物件資料刪除
        [HttpPost]
        public void DeletePickingItem(string PickingItemIds = "")
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.DeletePickingItem(PickingItemIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePickingLot 刪除已揀貨無條碼批號
        [HttpPost]
        public void DeletePickingLot(int PickingLotId = -1)
        {
            try
            {
                PickingApiLoginCheck();

                #region //Request
                pickingSystemDA = new PickingSystemDA();
                dataRequest = pickingSystemDA.DeletePickingLot(PickingLotId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Custom
        #region //登入
        [HttpPost]
        public void PickingLogin(string SystemKey, string Account)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetSubSystemLogin(SystemKey, Account, "PickingSystem");
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    LoginLog(Convert.ToInt32(result[0]["UserId"]), Account, "Picking");
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

        #region //View檢查登入
        [NonAction]
        public void PickingLoginCheck()
        {
            bool verify = LoginVerify("Picking");

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "PickingSystem");

                var uri = new Uri(redirectUrl);
                var newQueryString = HttpUtility.ParseQueryString(uri.Query);

                string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path).Replace("http://", "").Replace(Request.Url.Authority, "");
                redirectUrl = newQueryString.Count > 0 ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString) : pagePathWithoutQueryString;
                if (redirectUrl != "/") redirectLoginUrl += "?preUrl=" + Uri.EscapeDataString(redirectUrl);

                HttpContext.Response.Redirect(redirectLoginUrl);
            }
        }
        #endregion

        #region //Api檢查登入
        [NonAction]
        public void PickingApiLoginCheck()
        {
            bool verify = LoginVerify("Picking");

            if (!verify)
            {
                throw new SystemException("系統已登出，請重新登入");
            }
        }
        #endregion
        #endregion
    }
}