using ClosedXML.Excel;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NiceLabel.SDK;
using PDMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ProductController : WebController
    {
        private ProductDA productDA = new ProductDA();

        #region //View
        public ActionResult MtlItem()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult MtlItemChange()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult BomChange()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MtlItemTypeConstituteLogic()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult BomMtlItemPlatform()
        {
            ViewLoginCheck();
            return View();
        }

        /*20240807*/
        public ActionResult MassProductionReviewDoc()
        {
            ViewLoginCheck();

            return View();
        }

        /*20241108*/
        public ActionResult MtlItemNoIntegration()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetMtlItem 取得品號資料
        [HttpPost]
        public void GetMtlItem(int MtlItemId = -1, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", string CustomerMtlItemNo = "", string SearchKey = ""
            , string Status = "", string TransferStatus = "", string BomTransferStatus = "", string CustomerTransferStatus = "", string BomSubstitutionTransferStatus = "", string UomCTransferStatus = "", string CheckMtlDate = ""
            , int ItemTypeId = -1, string LotManagement= "", string ItemAttribute ="", string TransferStartDate = "", string TransferEndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
               
                WebApiLoginCheck("MtlItem", "read,constrained-data");
                productDA = new ProductDA();
                #region //Request
                dataRequest = productDA.GetMtlItem(MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, CustomerMtlItemNo, SearchKey
                    , Status, TransferStatus, BomTransferStatus, CustomerTransferStatus, BomSubstitutionTransferStatus, UomCTransferStatus, CheckMtlDate
                    , ItemTypeId, LotManagement, ItemAttribute, TransferStartDate, TransferEndDate
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

        #region //GetCustomerMtlItem 取得客戶品號資料
        [HttpPost]
        public void GetCustomerMtlItem(int CustomerMtlItemId = -1, int MtlItemId = -1, int CustomerId = -1, string CustomerMtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "customer-mtlItem");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetCustomerMtlItem(CustomerMtlItemId, MtlItemId, CustomerId, CustomerMtlItemNo
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

        #region //GetBillOfMaterials 取得Bom主件資料
        [HttpPost]
        public void GetBillOfMaterials(int BomId = -1, int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBillOfMaterials(BomId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomDetail 取得Bom元件資料
        [HttpPost]
        public void GetBomDetail(int BomDetailId = -1, int BomId = -1, int MtlItemId = -1,string HideFlag="")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomDetail(BomDetailId, BomId, MtlItemId, HideFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomSequence 取得Bom元件序號資料
        [HttpPost]
        public void GetBomSequence(int BomId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomSequence(BomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRdDesign 取得研發設計圖資料
        [HttpPost]
        public void GetRdDesign(int MtlItemId = -1, string CustomerMtlItemNo = "", string MtlItemIsNull = "", string ReleasedStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "design");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetRdDesign(MtlItemId, CustomerMtlItemNo, MtlItemIsNull, ReleasedStatus
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

        #region //GetBomSubstitution 取得取替代料資料
        [HttpPost]
        public void GetBomSubstitution(int MtlItemId = -1, int SubstitutionId = -1, int BomDetailId = -1, int SubMtlItemId = -1, string SubstituteStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "substitution");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomSubstitution(MtlItemId, SubstitutionId, BomDetailId, SubMtlItemId, SubstituteStatus
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

        #region //GetMtlItemSetting 取得品號設定資料
        [HttpPost]
        public void GetMtlItemSetting(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "setting");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlItemSetting(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlItemItNoSetteing 取得品號結構設定
        [HttpPost]
        public void GetMtlItemItNoSetteing(int MtlItemId = -1, int ItemTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "setting");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlItemItNoSetteing(MtlItemId, ItemTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCheckMtlItemName 判斷品名是否重複
        [HttpPost]
        public void GetCheckMtlItemName(string MtlItemName = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "setting");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetCheckMtlItemName(MtlItemName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMtlChangeEdition 查詢品號變更單目前版本
        [HttpPost]
        public void GetMtlChangeEdition(int MtlItemId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItemChange", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlChangeEdition(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBOMList --取得複製BOM用列表
        [HttpPost]
        public void GetBOMList(int CurrentMtlItemId = -1, string MtlItem = "", string TransferStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBOMList(CurrentMtlItemId, MtlItem, TransferStatus
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

        #region //GetMtlItemUomCalculate 取得品號單位換算
        [HttpPost]
        public void GetMtlItemUomCalculate(int MtlItemId = -1, int MtlItemUomCalculateId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlItemUomCalculate(MtlItemId, MtlItemUomCalculateId
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

        #region //GetMtlItemUomCalculateSimple 取得品號單位換算(下拉用)
        [HttpPost]
        public void GetMtlItemUomCalculateSimple(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlItemUomCalculateSimple(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlChangeData --取得品號變更列表
        [HttpPost]
        public void GetMtlChangeData(string SearchKey = "", string ChangeEdition = "", string TransferStatus = "", string ConfirmStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlChangeData(SearchKey, ChangeEdition, TransferStatus, ConfirmStatus, StartDate, EndDate
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

        #region //GetMtlChangeDetail --取得該筆品號變更
        [HttpPost]
        public void GetMtlChangeDetail(int MtlItemChangeId = -1, string ChangeEdition = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlChangeDetail(MtlItemChangeId, ChangeEdition);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlChangeDetailEdit --取得該筆品號變更
        [HttpPost]
        public void GetMtlChangeDetailEdit(int MtlItemChangeDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlChangeDetailEdit(MtlItemChangeDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomChangeData --取得BOM變更列表
        [HttpPost]
        public void GetBomChangeData(string SearchKey = "", string TransferStatus = "", string ConfirmStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomChangeData(SearchKey, TransferStatus, ConfirmStatus, StartDate, EndDate
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

        #region //GetBomChangeDetail --取得該筆BOM變更 GetBomChangeSubDetail
        [HttpPost]
        public void GetBomChangeDetail(int BomChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomChangeDetail(BomChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomChangeDetailEdit --取得該筆BOM變更單身 
        [HttpPost]
        public void GetBomChangeDetailEdit(int BomChangeDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomChangeDetailEdit(BomChangeDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomChangeSubDetailEdit --取得該筆BOM變更單身 
        [HttpPost]
        public void GetBomChangeSubDetailEdit(int BomChangeSubDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomChangeSubDetailEdit(BomChangeSubDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProduceLine 取得ERP生產線別
        [HttpPost]
        public void GetProduceLine(string ProduceLineNo = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetProduceLine(ProduceLineNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GePriceSix 取得ERP品號售價定價六
        [HttpPost]
        public void GePriceSix(string MtlitemNo = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GePriceSix(MtlitemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GeErpUnit 取得ERP單位
        [HttpPost]
        public void GeErpUnit()
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GeErpUnit();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPreday 取得ERP前置天數
        [HttpPost]
        public void GetPreday(string MtlitemNo = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetPreday(MtlitemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomChangeSubDetail --取得該筆BOM變更子單身
        [HttpPost]
        public void GetBomChangeSubDetail(int BomChangeDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomChangeSubDetail(BomChangeDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetWipPreFix --取得該筆BOM變更
        [HttpPost]
        public void GetWipPreFix(string WipPrefix)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetWipPreFix(WipPrefix);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlItemType --取得類別
        [HttpPost]
        public void GetMtlItemType(string Condition = "", string TypeNo = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlItemType(Condition, TypeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBomChangeErpNo --取得BOM變更單Erp編號
        [HttpPost]
        public void GetBomChangeErpNo()
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetBomChangeErpNo();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlQcItem 取得品號進出貨檢資料
        [HttpPost]
        public void GetMtlQcItem(int MtlItemId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "read,constrained-data");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlQcItem(MtlItemId
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

        #region //GeMtlItemTypeLogic -- 取得品號類型編碼邏輯
        [HttpPost]
        public void GeMtlItemTypeLogic(string MtlItemType = "", string MtlItemTypeNo = "", string MtlItemTypeName = "", string MtlItemNameDesc = "", string FirstWord = "", string SecondWord = "", string ThirdWord = "", string FourthWord = "", string FifthWord = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItemTypeConstituteLogic", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GeMtlItemTypeLogic(MtlItemType, MtlItemTypeNo, MtlItemTypeName, MtlItemNameDesc, FirstWord, SecondWord, ThirdWord, FourthWord, FifthWord
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

        #region //GeMtlItemType -- 取得品號類型編碼邏輯
        [HttpPost]
        public void GeMtlItemType()
        {
            try
            {
                WebApiLoginCheck("MtlItemTypeConstituteLogic", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GeMtlItemType();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //GeMtlItemTypeWordNode -- 取得品號類型節點
        [HttpPost]
        public void GeMtlItemTypeWordNode(int SortNumber = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItemTypeConstituteLogic", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GeMtlItemTypeWordNode(SortNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        /*20240807*/
        #region //GetMassProductionReviewDocSetting -- 取得量產審查文件設定
        [HttpPost]
        public void GetMassProductionReviewDocSetting(int DocId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMassProductionReviewDocSetting(DocId, Status
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

        /*20240807*/
        #region //GetMassProductionReviewDoc -- 取得量產審查文件
        [HttpPost]
        public void GetMassProductionReviewDoc(int DocId = -1, string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMassProductionReviewDoc(DocId, MtlItemNo, MtlItemName
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

        #region //GetInventoryQueryWithNewMtlItemNo 庫存整合+新舊品號變更查詢
        [HttpPost]
        public void GetInventoryQueryWithNewMtlItemNo(
            string SearchKey = "",
            string SearchMtlItemNo = "",
            string SearchNewMtlItemNo = "",
            string SearchMtlItemSpec = "",
            string UserIds = "",
            string NewUserIds = "",
            string InStock = "",
            string UserStatus = "",
            string NewUserStatus = "",
            string OrderBy = "",
            int PageIndex = -1,
            int PageSize = -1
        )
        {
            try
            {
                WebApiLoginCheck("MtlItemNoIntegration", "read");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetInventoryQueryWithNewMtlItemNo(
                    SearchKey,
                    SearchMtlItemNo,
                    SearchNewMtlItemNo,
                    SearchMtlItemSpec,
                    UserIds,
                    NewUserIds,
                    InStock,
                    UserStatus,
                    NewUserStatus,
                    OrderBy,
                    PageIndex,
                    PageSize
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

        #region //GetMtlItemFile -- 取得品號審查文件
        [HttpPost]
        public void GetMtlItemFile(int MtlItemId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetMtlItemFile(MtlItemId
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

        #region //GetNewMtlItemDetail -- 取得新品號基本資料
        [HttpPost]
        public void GetNewMtlItemDetail(int MtlItemId = -1, string NewMtlItemNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetNewMtlItemDetail(MtlItemId, NewMtlItemNo
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

        #region //GetActionSituationSimple -- 取得新品號當前操作情境
        [HttpPost]
        public void GetActionSituationSimple(string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.GetActionSituationSimple(MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddMtlItem 品號資料新增
        [HttpPost]
        public void AddMtlItem(int MtlItemId = -1, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", int InventoryUomId = -1, int PurchaseUomId = -1, int SaleUomId = -1
            , string TypeOne = "", string TypeTwo = "", string TypeThree = "", string TypeFour = "", int InventoryId = -1, int RequisitionInventoryId = -1, string InventoryManagement = "", string MtlModify = ""
            , string BondedStore = "", string ItemAttribute = "", string LotManagement = "", string MeasureType = "", string OverReceiptManagement = "", string OverDeliveryManagement = "", string EffectiveDate = ""
            , string ExpirationDate = "", string MtlItemDesc = "", string MtlItemRemark = "", string CustomerMtlItems = "", int ItemTypeId =-1, string ItemTypeSegmentList = "", string ItemTypeEnable = ""
            , string AddBillOfMaterials = "", int EfficientDays = -1, int RetestDays = -1, string ReplenishmentPolicy = "", string ProductionLine ="")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddMtlItem(MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, InventoryUomId, PurchaseUomId, SaleUomId
                    , TypeOne, TypeTwo, TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryManagement, MtlModify, BondedStore, ItemAttribute
                    , LotManagement, MeasureType, OverReceiptManagement, OverDeliveryManagement, EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark, CustomerMtlItems, ItemTypeId, ItemTypeSegmentList, ItemTypeEnable
                    , AddBillOfMaterials, EfficientDays, RetestDays, ReplenishmentPolicy, ProductionLine);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMtlItemBatch 品號資料新增(批量)
        [HttpPost]
        public void AddMtlItemBatch(string MtlItemData = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "batch-add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddMtlItemBatch(MtlItemData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBomDetail Bom元件資料新增
        [HttpPost]
        public void AddBomDetail(int BomId = -1, int MtlItemId = -1, string BomSequence = "", double CompositionQuantity = 0.0, double Base = 0.0
            , double LossRate = 0.0,string StandardCostingType="",string  MaterialProperties="", string EffectiveDate = "", string ExpirationDate = "", string Remark = "", string ComponentType = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddBomDetail(BomId, MtlItemId, BomSequence, CompositionQuantity, Base, LossRate, StandardCostingType, MaterialProperties, EffectiveDate, ExpirationDate, Remark, ComponentType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBomSubstitution 取替代料資料新增
        [HttpPost]
        public void AddBomSubstitution(int MainMtlItemId = -1, int BomDetailId = -1, string SubstituteStatus = "", int MtlItemId = -1, double Quantity = 0.0
            , int SortNumber = -1, string Precedence = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "substitution");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddBomSubstitution(MainMtlItemId, BomDetailId, SubstituteStatus, MtlItemId, Quantity, SortNumber, Precedence, EffectiveDate, ExpirationDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddCustomerMtlItem 部番資料新增
        [HttpPost]
        public void AddCustomerMtlItem(int MtlItemId = -1, int CustomerId = -1, string CustomerMtlItemNo = "", string CustomerMtlItemName = "", string CustomerMtlItemSpec = "", int PoCompanyId = -1, string PoErpPrefixNo = "", string PoSeq = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "customer-mtlItem");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddCustomerMtlItem(MtlItemId, CustomerId, CustomerMtlItemNo, CustomerMtlItemName, CustomerMtlItemSpec, PoCompanyId, PoErpPrefixNo, PoSeq);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddCopyMtlItem 複製品號
        [HttpPost]
        public void AddCopyMtlItem(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "copy-mtlitem");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddCopyMtlItem(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMtlItemChange 品號變更資料新增
        [HttpPost]
        public void AddMtlItemChange(int MtlItemId = -1, string OriginalMtlItemName = "", string OriginalMtlItemSpec = ""
            , string ChangeEdition = "", string ChangeDate = "", string DelMtlItemStatus = "", string NewMtlItemName = ""
            , string NewMtlItemSpec = ""
            , int UpdateBy = -1, string Remark = "", string FieldType = "", int ConfirmBy = -1, string ConfirmDate = ""
            , string ConfirmStatus = "", string SignOffStatus = "", string PrintCount = ""
            , string SendCount = "", string OriginalEdition = "", string TransferStatus = "", string TransferDate = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItemChange", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddMtlItemChange(MtlItemId, ChangeEdition, OriginalMtlItemName, OriginalMtlItemSpec, ChangeDate, DelMtlItemStatus, NewMtlItemName, NewMtlItemSpec
                    , UpdateBy, Remark, FieldType, ConfirmBy, ConfirmDate, ConfirmStatus, SignOffStatus, PrintCount
                    , SendCount, OriginalEdition, TransferStatus, TransferDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion 

        #region //AddMtlItemChangeDetail 品號變更欄位資料新增
        [HttpPost]
        public void AddMtlItemChangeDetail(int MtlItemChangeId = -1, int MtlItemId = -1, string ChangeEdition = "", string FieldNum = "", string FieldName = ""
            , string NewTextField = "", string NewNumericField = "", string OriginalNumericField = "", string NewFieldName = "", string OriginalFieldName = "", string Remark = ""
            , string FieldType = "", string RetentionFied = "", string Status = "", string SortNum = "", string OriginalTextField = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItemChange", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddMtlItemChangeDetail(MtlItemChangeId, MtlItemId, ChangeEdition, FieldNum, FieldName, NewTextField, NewNumericField
                    , OriginalNumericField, NewFieldName, OriginalFieldName, Remark, FieldType, RetentionFied, Status, SortNum, OriginalTextField);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion 

        #region //AddCopyBomDetail Bom元件資料複製
        [HttpPost]
        public void AddCopyBomDetail(int CurrentBomId = -1, int BomId = -1, int MtlItemId = -1, int bomSequenceVal = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddCopyBomDetail(CurrentBomId, MtlItemId, BomId, bomSequenceVal);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMtlItemUomCalculate 品號單位換算新增
        [HttpPost]
        public void AddMtlItemUomCalculate(int MtlItemId = -1, int UomId = -1, int ConvertUomId = -1, double SwapNumerator = -1, double SwapDenominator = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddMtlItemUomCalculate(MtlItemId, UomId, ConvertUomId, SwapNumerator, SwapDenominator);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBomChange BOM變更欄位資料新增
        [HttpPost]
        public void AddBomChange( string BomChangePrefix = "", string EmergencyCode = "", string ChangeReason = "", string Remark = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItemChange", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddBomChange(BomChangePrefix, EmergencyCode, ChangeReason, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBomChangeDetail BOM變更單身資料新增
        [HttpPost]
        public void AddBomChangeDetail(int BomChangeId = -1, int MtlItemId = -1, string Unit = "", string SmallUnit = ""
            , string StandardLot = "", string MoPrefix = "", string ChangeReason = "", string NewRemark = "", string ConfirmCode = ""
            , string ConfirmStatus = "", string SortNum = "", int BomId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItemChange", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddBomChangeDetail(BomChangeId, MtlItemId, Unit, SmallUnit, StandardLot, MoPrefix
                    , ChangeReason, ConfirmStatus, ConfirmCode, NewRemark, SortNum, BomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBomChangeSubDetail BOM變更子單身資料新增
        [HttpPost]
        public void AddBomChangeSubDetail(int BomChangeId = -1, int BomChangeDetailId = -1, int BomDetailId = -1, int MtlItemId = -1, int Unit = -1, double CompositionQuantity = -1, double Base = -1, string ComponentType = ""
            , double LossRate = -1, string EffectiveDate = "", string ExpirationDate = "", string ChangeReason = "", string NewRemark = "", int SortNum = 0, string Provide = "", string isChangeToNew = "", string StandardCostingType = "Y", string MaterialProperties = "1", string DeleteStatus = "", int BomId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItemChange", "add");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.AddBomChangeSubDetail(BomChangeId, BomChangeDetailId, BomDetailId, MtlItemId, Unit, CompositionQuantity, Base, ComponentType
            , LossRate, EffectiveDate, ExpirationDate, ChangeReason, NewRemark, SortNum, Provide, isChangeToNew, StandardCostingType, MaterialProperties, DeleteStatus, BomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddMtlItemFile --新增品號審核檔案
        [HttpPost]
        public void AddMtlItemFile(string MtlItemNo = "", int MtDocId = -1, string FileId = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "file-add");

                #region //Request 
                productDA = new ProductDA();
                dataRequest = productDA.AddMtlItemFile(MtlItemNo, MtDocId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateMtlItem 品號資料更新
        [HttpPost]
        public void UpdateMtlItem(int MtlItemId = -1, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", int InventoryUomId = -1, int PurchaseUomId = -1, int SaleUomId = -1
            , string TypeOne = "", string TypeTwo = "", string TypeThree = "", string TypeFour = "", int InventoryId = -1, int RequisitionInventoryId = -1, string InventoryManagement = "", string MtlModify = ""
            , string BondedStore = "", string ItemAttribute = "", string LotManagement = "", string MeasureType = "", string OverReceiptManagement = "", string OverDeliveryManagement = "", string EffectiveDate = ""
            , string ExpirationDate = "", string MtlItemDesc = "", string MtlItemRemark = "", string CustomerMtlItems = "", int EfficientDays = -1, int RetestDays = -1, string ProductionLine = ""
            , int ItemTypeId = -1, string ItemTypeSegmentList = "", string ItemTypeEnable = "",string ReplenishmentPolicy="")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItem(MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, InventoryUomId, PurchaseUomId, SaleUomId
                    , TypeOne, TypeTwo, TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryManagement, MtlModify, BondedStore, ItemAttribute
                    , LotManagement, MeasureType, OverReceiptManagement, OverDeliveryManagement, EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark, CustomerMtlItems
                    , EfficientDays, RetestDays, ProductionLine, ItemTypeId, ItemTypeSegmentList, ItemTypeEnable, ReplenishmentPolicy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemStatus 品號狀態更新
        [HttpPost]
        public void UpdateMtlItemStatus(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "status-switch");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItemStatus(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBillOfMaterials Bom主件資料更新
        [HttpPost]
        public void UpdateBillOfMaterials(int BomId = -1, int MtlItemId = -1, double StandardLot = 0.0, string WipPrefix = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBillOfMaterials(BomId, MtlItemId, StandardLot, WipPrefix, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomDetail Bom元件資料更新
        [HttpPost]
        public void UpdateBomDetail(int BomDetailId = -1, int BomId = -1, int MtlItemId = -1, string BomSequence = "", double CompositionQuantity = 0.0, double Base = 0.0, double LossRate = 0.0
            , string StandardCostingType = "", string MaterialProperties = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "", string ComponentType = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomDetail(BomDetailId, BomId, MtlItemId, BomSequence, CompositionQuantity, Base, LossRate
                    , StandardCostingType, MaterialProperties, EffectiveDate, ExpirationDate, Remark, ComponentType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateConfirmBomExcel 上傳BomExce資料
        [HttpPost]
        public void UpdateConfirmBomExcel(int BomId = -1, string BomExcelJson = "", string BomExcelData = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateConfirmBomExcel(BomId, BomExcelJson, BomExcelData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlRdDesign 研發設計圖品號更新
        [HttpPost]
        public void UpdateMtlRdDesign(int MtlItemId = -1, int DesignId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "design");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlRdDesign(MtlItemId, DesignId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlSaleOrder 訂單品號更新
        [HttpPost]
        public void UpdateMtlSaleOrder(int MtlItemId = -1, int SoDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "sale-order");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlSaleOrder(MtlItemId, SoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomSubstitution 取替代料資料更新
        [HttpPost]
        public void UpdateBomSubstitution(int SubstitutionId = -1, int BomDetailId = -1, string SubstituteStatus = "", int MtlItemId = -1, double Quantity = 0.0
            , int SortNumber = -1, string Precedence = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "substitution");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomSubstitution(SubstitutionId, BomDetailId, SubstituteStatus, MtlItemId, Quantity, SortNumber, Precedence, EffectiveDate, ExpirationDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSubstitutionSort 取替代料順序調整
        [HttpPost]
        public void UpdateSubstitutionSort(int BomDetailId = -1, string SubstitutionList = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "substitution");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateSubstitutionSort(BomDetailId, SubstitutionList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemToERP 品號資料拋轉
        [HttpPost]
        public void UpdateMtlItemToERP(int MtlItemId = -1)
        {
            string transferRequest = string.Empty;
            try
            {
                WebApiLoginCheck("MtlItem", "data-transfer");

                #region //Request 品號拋轉
                transferRequest = productDA.UpdateMtlItemToERP(MtlItemId);
                #endregion

                #region //Request 品號拋轉(停用)
                //if (JObject.Parse(transferRequest)["status"].ToString() == "success" || JObject.Parse(transferRequest)["route"].ToString() == "batch")
                //{
                //    #region //Request Bom拋轉
                //    transferRequest = productDA.UpdateBomToERP(MtlItemId, -1);
                //    #endregion
                //}

                //if (JObject.Parse(transferRequest)["status"].ToString() == "success" || JObject.Parse(transferRequest)["route"].ToString() == "batch")
                //{
                //    #region //Request 部番資料拋轉
                //    transferRequest = productDA.UpdateCustomerMtlItemToERP(MtlItemId);
                //    #endregion
                //}

                //if (JObject.Parse(transferRequest)["status"].ToString() == "success" || JObject.Parse(transferRequest)["route"].ToString() == "batch")
                //{
                //    #region //Request 取替代料拋轉
                //    transferRequest = productDA.UpdateBomSubstitutionToERP(MtlItemId);
                //    #endregion
                //}

                //if (JObject.Parse(transferRequest)["status"].ToString() == "success" || JObject.Parse(transferRequest)["route"].ToString() == "batch")
                //{
                //    #region //Request 品號單位換算拋轉(停用) 2023.09.11
                //    transferRequest = productDA.UpdateMtlItemUomCalculateToERP(MtlItemId);
                //    #endregion
                //}

                //var JObj = JObject.Parse(transferRequest);
                //if (JObj["route"].ToString() == "batch")
                //{
                //    JObj["status"] = "success";
                //    transferRequest = JObj.ToString();
                //}
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(transferRequest.ToString());
        }
        #endregion

        #region //UpdateBomToERP Bom資料拋轉
        [HttpPost]
        public void UpdateBomToERP(int MtlItemId = -1, int BomId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomToERP(MtlItemId, BomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCustomerMtlItemToERP 客戶品號拋轉ERP
        [HttpPost]
        public void UpdateCustomerMtlItemToERP(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "data-transfer");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateCustomerMtlItemToERP(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomSubstitutionToERP 取替代料拋轉
        [HttpPost]
        public void UpdateBomSubstitutionToERP(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "data-transfer");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomSubstitutionToERP(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCustomerMtlItem 部番資料更新
        [HttpPost]
        public void UpdateCustomerMtlItem(int CustomerMtlItemId = -1, int MtlItemId = -1, int CustomerId = -1, string CustomerMtlItemNo = "", string CustomerMtlItemName = "", string CustomerMtlItemSpec = "", int PoCompanyId = -1, string PoErpPrefixNo = "", string PoSeq = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "customer-mtlItem");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateCustomerMtlItem(CustomerMtlItemId, MtlItemId, CustomerId, CustomerMtlItemNo, CustomerMtlItemName, CustomerMtlItemSpec, PoCompanyId, PoErpPrefixNo, PoSeq);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemSetting 品號設定資料更新
        [HttpPost]
        public void UpdateMtlItemSetting(int MtlItemId = -1, int FixedLeadTime = -1, int ChangeLeadTime = -1, int MinPoQty = -1, int MultiplePoQty = -1, int MultipleReqQty = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "setting");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItemSetting(MtlItemId, FixedLeadTime, ChangeLeadTime, MinPoQty, MultiplePoQty, MultipleReqQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemUomCalculate 品號單位換算更新
        [HttpPost]
        public void UpdateMtlItemUomCalculate(int MtlItemUomCalculateId = -1, int MtlItemId = -1, int ConvertUomId = -1, double SwapNumerator = -1, double SwapDenominator = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItemUomCalculate(MtlItemUomCalculateId, MtlItemId, ConvertUomId, SwapNumerator, SwapDenominator);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemUomCalculateToERP 品號單位換算拋轉
        [HttpPost]
        public void UpdateMtlItemUomCalculateToERP(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "data-transfer");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItemUomCalculateToERP(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemChange -- 品號變更單頭資料更新
        [HttpPost]
        public void UpdateMtlItemChange(int MtlItemChangeId = -1, int MtlItemId = -1, string NewMtlItemName = "", string NewMtlItemSpec = "", string Remark = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItemChange(MtlItemChangeId, MtlItemId, NewMtlItemName, NewMtlItemSpec, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlItemChangeDetail -- 品號變更單身資料更新
        [HttpPost]
        public void UpdateMtlItemChangeDetail(int MtlItemChangeId = -1, int MtlChangeDetailId = -1, string FieldNum = "", string FieldName = "", string NewTextField = "", string NewNumericField = "", string OriginalNumericField = "", string NewFieldName = "", string OriginalFieldName = "", string OriginalTextField = ""
            , string Remark = "", string FieldType = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlItemChangeDetail(MtlItemChangeId, MtlChangeDetailId, FieldNum, FieldName, NewTextField, NewNumericField, OriginalNumericField, NewFieldName, OriginalFieldName, OriginalTextField, Remark, FieldType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChange -- BOM變更單頭資料更新 
        [HttpPost]
        public void UpdateBomChange(int BomChangeId = -1, string EmergencyCode = "", string ChangeReason = "", string Remark = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChange(BomChangeId, EmergencyCode, ChangeReason, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChangeDetail -- 修改Bom變更單身 
        [HttpPost]
        public void UpdateBomChangeDetail(int BomChangeDetailId = -1, string MoPrefix = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChangeDetail(BomChangeDetailId, MoPrefix);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChangeSubDetail -- 修改Bom變更子單身 
        [HttpPost]
        public void UpdateBomChangeSubDetail(int BomChangeSubDetailId = -1, double CompositionQuantity = -1, double Base = -1, string ComponentType = ""
            , double LossRate = -1, string EffectiveDate = "", string ExpirationDate = "", string ChangeReason = "", string NewRemark = "", string StandardCostingType = "Y", string MaterialProperties = "1")
      
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChangeSubDetail(BomChangeSubDetailId, CompositionQuantity, Base, ComponentType, LossRate, EffectiveDate, ExpirationDate, ChangeReason, NewRemark, StandardCostingType, MaterialProperties);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlitemChangeToErp --品號變更單拋轉Erp
        [HttpPost]
        public void UpdateMtlitemChangeToErp(int MtlItemChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlitemChangeToErp(MtlItemChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChangeToErp --Bom變更單拋轉Erp
        [HttpPost]
        public void UpdateBomChangeToErp(int BomChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChangeToErp(BomChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtiitemChangeConfirm --品號變更單確認
        [HttpPost]
        public void UpdateMtiitemChangeConfirm(int MtlItemChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtiitemChangeConfirm(MtlItemChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtiitemChangeReConfirm --品號變更單反確認
        [HttpPost]
        public void UpdateMtiitemChangeReConfirm(int MtlItemChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtiitemChangeReConfirm(MtlItemChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtiitemChangeVoid --品號變更單作廢
        [HttpPost]
        public void UpdateMtiitemChangeVoid(int MtlItemChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtiitemChangeVoid(MtlItemChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChangeConfirm --BOM變更單確認
        [HttpPost]
        public void UpdateBomChangeConfirm(int BomChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChangeConfirm(BomChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChangeReConfirm //BOM變更單反確認
        [HttpPost]
        public void UpdateBomChangeReConfirm(int BomChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChangeReConfirm(BomChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBomChangeVoid //BOM變更單作廢
        [HttpPost]
        public void UpdateBomChangeVoid(int BomChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateBomChangeVoid(BomChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlQcItemBatch 品號進出貨檢資料更新
        [HttpPost]
        public void UpdateMtlQcItemBatch(int MtlItemId = -1, string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMtlQcItemBatch(MtlItemId, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        /*20240807*/
        #region //UpdateMassProductionReviewDocFile 量產審查文件檔案更新
        [HttpPost]
        public void UpdateMassProductionReviewDocFile(int MtlItemId = -1, int DocId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("MassProductionReviewDoc", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateMassProductionReviewDocFile(MtlItemId, DocId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateInventoryQueryWithNewMtlItemNo 新舊品號更新
        public void UpdateInventoryQueryWithNewMtlItemNo(
            int MtlItemId = -1,
            string NewMtlItemNo = "",
            string NewMtlItemName = "",
            string NewMtlItemSpec = ""
        )
        {
            try
            {
                WebApiLoginCheck("MtlItemNoIntegration", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateInventoryQueryWithNewMtlItemNo(MtlItemId, NewMtlItemNo, NewMtlItemName, NewMtlItemSpec);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateNewMtlItem 品號資料更新
        [HttpPost]
        public void UpdateNewMtlItem(int MtlItemId, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", int InventoryUomId = -1, int PurchaseUomId = -1, int SaleUomId = -1
            , string TypeOne = "", string TypeTwo = "", string TypeThree = "", string TypeFour = "", int InventoryId = -1, int RequisitionInventoryId = -1, string InventoryManagement = "", string MtlModify = ""
            , string BondedStore = "", string ItemAttribute = "", string LotManagement = "", string MeasureType = "", string OverReceiptManagement = "", string OverDeliveryManagement = "", string EffectiveDate = ""
            , string ExpirationDate = "", string MtlItemDesc = "", string MtlItemRemark = "", int EfficientDays = -1, int RetestDays = -1, string ProductionLine = "", string ReplenishmentPolicy = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateNewMtlItem(MtlItemId,MtlItemNo, MtlItemName, MtlItemSpec, InventoryUomId, PurchaseUomId, SaleUomId
                    , TypeOne, TypeTwo, TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryManagement, MtlModify, BondedStore, ItemAttribute
                    , LotManagement, MeasureType, OverReceiptManagement, OverDeliveryManagement, EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark
                    , EfficientDays, RetestDays, ProductionLine, ReplenishmentPolicy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateNewMtlItemToErp 新舊品號更新
        public void UpdateNewMtlItemToErp(string MtlItemList = "")
        {
            try
            {
                WebApiLoginCheck("MtlItemNoIntegration", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateNewMtlItemToErp(MtlItemList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #region //Excel
                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式

                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
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
                    string excelFileName = "品號異動歷程記錄";
                    string [] excelsheetName =new[] { "新品號拋轉動態", "新舊庫存調整動態", "過程中異常" };

                    var jsonString = JObject.Parse(dataRequest)["data"].ToString();
                    LogData logData = JsonConvert.DeserializeObject<LogData>(jsonString);
                    string[] header = new string[] { "舊品號","新品號","回報記錄"};
                    string colIndex = "";
                    int sheetIndex = 0;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        foreach (var sheetName in excelsheetName)
                        {
                            var worksheet = workbook.Worksheets.Add(sheetName);
                            worksheet.RowHeight = 16;
                            worksheet.Style = defaultStyle;
                            int rowIndex = 1;

                            #region //HEADER
                            for (int i = 0; i < header.Length; i++)
                            {
                                colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                worksheet.Cell(colIndex).Value = header[i];
                                worksheet.Cell(colIndex).Style = headerStyle;
                            }
                            #endregion

                            #region //BODY
                            if(sheetIndex == 0)
                            {
                                foreach (var transferLog in logData.TransferLogs)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = transferLog.Split(':')[0];
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = transferLog.Split(':')[1];
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = transferLog.Split(':')[2];
                                }
                            }
                            if(sheetIndex == 1)
                            {
                                foreach (var adjustmentLog in logData.AdjustmentLogs)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = adjustmentLog.Split(':')[0];
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = adjustmentLog.Split(':')[1];
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = adjustmentLog.Split(':')[2];
                                }
                            }
                            if (sheetIndex == 2)
                            {
                                foreach (var errLog in logData.ErrLogs)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = errLog.Split(':')[0];
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = errLog.Split(':')[1];
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = errLog.Split(':')[2];
                                }
                            }
                            #endregion

                            #region //自適應欄寬
                            worksheet.Columns().AdjustToContents();
                            #endregion

                            sheetIndex++;
                        }

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
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        public class LogData
        {
            public List<string> TransferLogs { get; set; }
            public List<string> AdjustmentLogs { get; set; }
            public List<string> ErrLogs { get; set; }
        }

        #endregion

        #region //UpdateInventoryTransactionToErp 新舊品號異動單拋轉ERP
        public void UpdateInventoryTransactionToErp(string MtlItemList = "")
        {
            try
            {
                WebApiLoginCheck("MtlItemNoIntegration", "update");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.UpdateInventoryTransactionToErp(MtlItemList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteMtlItem 品號資料刪除
        [HttpPost]
        public void DeleteMtlItem(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlItem(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBomDetail Bom元件資料刪除
        [HttpPost]
        public void DeleteBomDetail(int BomDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteBomDetail(BomDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBomDetailAll 刪除Bom元件全部
        [HttpPost]
        public void DeleteBomDetailAll(int BomId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteBomDetailAll(BomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMtlRdDesign 研發設計圖品號刪除
        [HttpPost]
        public void DeleteMtlRdDesign(int DesignId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "design");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlRdDesign(DesignId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMtlSaleOrder 訂單品號刪除
        [HttpPost]
        public void DeleteMtlSaleOrder(int SoDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "sale-order");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlSaleOrder(SoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBomSubstitution 取替代料資料刪除
        [HttpPost]
        public void DeleteBomSubstitution(int SubstitutionId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "substitution");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteBomSubstitution(SubstitutionId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteCustomerMtlItem 部番資料刪除
        [HttpPost]
        public void DeleteCustomerMtlItem(int CustomerMtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "customer-mtlItem");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteCustomerMtlItem(CustomerMtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMtlItemUomCalculate 品號單位換算刪除
        [HttpPost]
        public void DeleteMtlItemUomCalculate(int MtlItemUomCalculateId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlItemUomCalculate(MtlItemUomCalculateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteMtlItemChange --品號變更單投刪除
        [HttpPost]
        public void DeleteMtlItemChange(int MtlItemChangeId = -1, int MtlItemId = -1, string ChangeEdition = "")
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlItemChange(MtlItemChangeId, MtlItemId, ChangeEdition);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteMtlItemChangeDetail --品號變更單身刪除
        [HttpPost]
        public void DeleteMtlItemChangeDetail(int MtlChangeDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlItemChangeDetail(MtlChangeDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteBomChange --BOM變更單投刪除
        [HttpPost]
        public void DeleteBomChange(int BomChangeId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteBomChange(BomChangeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteBomChangeDetail --BOM變更單投刪除
        [HttpPost]
        public void DeleteBomChangeDetail(int BomChangeDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteBomChangeDetail(BomChangeDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteBomChangeSubDetail --BOM變更單投刪除
        [HttpPost]
        public void DeleteBomChangeSubDetail(int BomChangeDetailId = -1, int BomChangeSubDetailId = -1)
        {
            try
            {
                //WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteBomChangeSubDetail(BomChangeDetailId, BomChangeSubDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMtlQcItem --刪除品號進出貨檢資料刪除 -- Chia Yuan 2024.03.27
        [HttpPost]
        public void DeleteMtlQcItem(int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteMtlQcItem(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateMtlItemSynchronize -- 品號資料同步
        [HttpPost]
        [Route("api/ERP/MtlItemSynchronize")]
        public void UpdateMtlItemSynchronize(string Company, string SecretKey, string UpdateDate, string MtlItemNo)
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "UpdateMtlItemSynchronize");
                #endregion

                #region //Request
                dataRequest = productDA.UpdateMtlItemSynchronize(Company, UpdateDate, MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteMtlItemSynchronize -- 品號資料異動
        [HttpPost]
        [Route("api/ERP/DeleteMtlItemSynchronize")]
        public void DeleteMtlItemSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                int run = 0;
                if (run == 1) {
                    #region //Api金鑰驗證
                    ApiKeyVerify(Company, SecretKey, "UpdateMtlItemSynchronize");
                    #endregion

                    #region //Request
                    dataRequest = productDA.DeleteMtlItemSynchronize(Company, UpdateDate);
                    #endregion
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

        #region //UpdateBomSynchronize -- BOM資料同步
        [HttpPost]
        [Route("api/ERP/BomSynchronize")]
        public void UpdateBomSynchronize(string Company, string SecretKey, string UpdateDate, string MtlItemNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateBomSynchronize");
                #endregion

                #region //Request
                dataRequest = productDA.UpdateBomSynchronize(Company, UpdateDate, MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteBomSynchronize -- BOM資料異動
        [HttpPost]
        [Route("api/ERP/DeleteBomSynchronize")]
        public void DeleteBomSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                int run = 0;
                if (run == 1)
                {
                    #region //Api金鑰驗證
                    ApiKeyVerify(Company, SecretKey, "UpdateBomSynchronize");
                    #endregion

                    #region //Request
                    dataRequest = productDA.DeleteBomSynchronize(Company, UpdateDate);
                    #endregion
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

        #region //UpdateBomSubstitutionSynchronize -- BOM取替代料資料同步
        [HttpPost]
        [Route("api/ERP/BomSubstitutionSynchronize")]
        public void UpdateBomSubstitutionSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateBomSubstitutionSynchronize");
                #endregion

                #region //Request
                dataRequest = productDA.UpdateBomSubstitutionSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateConfrimSynchronize 品號資料同步(單一) (停用)
        //[HttpPost]
        //public void UpdateConfrimSynchronize(string MtlItemNo = "")
        //{
        //    try
        //    {
        //        WebApiLoginCheck("MtlItem", "update");

        //        #region //Request
        //        dataRequest = productDA.UpdateConfrimSynchronize(MtlItemNo);
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

        #region //UpdateMtlItemUomCalculateSynchronize -- 品號單位換算資料同步
        [HttpPost]
        [Route("api/ERP/UpdateMtlItemUomCalculateSynchronize")]
        public void UpdateMtlItemUomCalculateSynchronize(string Company, string UpdateDate)
        {
            try
            {
                //#region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "UpdateMtlItemSynchronize");
                //#endregion

                #region //Request
                dataRequest = productDA.UpdateMtlItemUomCalculateSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteCorrectMtlItem 品號異動刪除
        [HttpPost]
        public void DeleteCorrectMtlItem(string Type = "", string MtlItemNo = "", string CheckTableList = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "update");

                #region //Request
                dataRequest = productDA.DeleteCorrectMtlItem(Type, MtlItemNo, CheckTableList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteCorrectBom Bom表異動刪除
        [HttpPost]
        public void DeleteCorrectBom(string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom-delete");

                #region //Request
                productDA = new ProductDA();
                dataRequest = productDA.DeleteCorrectBom(MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //PrintLabelMtlItem -- 製令標籤列印  -- Shintokuro 2024-11-04
        [HttpPost]
        public void PrintLabelMtlItem(string MtlItemList = "", string PrintMachine = "", string DeliveryDate = "", int PrintNum = -1, string ViewCompanyId = "")
        {
            try
            {
                WebApiLoginCheck("LotBarcoderPrintLabel", "print");

                if (MtlItemList.Length <= 0) throw new SystemException("品號列表不可以為空,請重新確認!!");
                if (PrintMachine.Length <= 0) throw new SystemException("標籤機不可以為空,請重新確認!!");
                if (DeliveryDate.Length <= 0) throw new SystemException("交期不可以為空,請重新確認!!");
                if (PrintNum <= 0) throw new SystemException("列印數量不可以為空,請重新確認!!");

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
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintMtlItem\\Template\\LabelName.txt", Encoding.Default);
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
                dataRequest = productDA.GetMtlItemList(MtlItemList);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #region //資料操作
                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    for (int n = 0; n < PrintNum; n++)
                    {
                        foreach (var item in result)
                        {
                            #region //製令資料取值
                            string MtlItemNo = item["MtlItemNo"].ToString();
                            string MtlItemName = item["MtlItemName"].ToString();
                            #endregion

                            #region //標籤檔賦值
                            label.Variables["MtlItemNo"].SetValue(MtlItemNo);
                            label.Variables["MtlItemName"].SetValue(MtlItemName);
                            label.Variables["DeliveryDate"].SetValue(DeliveryDate);
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
        #endregion

        #region //Download

        #region //ExcelMtlItem 品號輸出Excel
        public void ExcelMtlItem(int MtlItemId = -1, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", string CustomerMtlItemNo = "", string SearchKey = ""
            , string Status = "", string TransferStatus = "", string BomTransferStatus = "", string CustomerTransferStatus = "", string BomSubstitutionTransferStatus = "", string UomCTransferStatus = "", string CheckMtlDate = ""
            , int ItemTypeId = -1, string LotManagement = "", string ItemAttribute = "", string TransferStartDate = "", string TransferEndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "read,constrained-data");
                List<string> OsrIdList = new List<string>();

                #region //Request
                dataRequest = productDA.GetMtlItem(MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, CustomerMtlItemNo, SearchKey
                    , Status, TransferStatus, BomTransferStatus, CustomerTransferStatus, BomSubstitutionTransferStatus, UomCTransferStatus, CheckMtlDate
                    , ItemTypeId, LotManagement, ItemAttribute, TransferStartDate, TransferEndDate
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
                    string excelFileName = "【MES2.0】品號";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "品號", "品名", "規格"};
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;


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
                            string mtlItemNo = item.MtlItemNo.ToString();
                            string mtlItemName = item.MtlItemName.ToString();
                            string mtlItemSpec = item.MtlItemSpec.ToString();
                            
                            var resultList = new[]
                            {
                                new {
                                    mtlItemNo,
                                    mtlItemName,
                                    mtlItemSpec,
                                }
                            };

                            rowIndex++;

                            var item1 = resultList[0]; // 因為你目前陣列只有一筆

                            var properties = item1.GetType().GetProperties(); // 用 reflection 取得屬性
                            for (int i = 0; i < properties.Length; i++)
                            {
                                var value = properties[i].GetValue(item1)?.ToString() ?? "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, rowIndex)).Value = value;
                            }
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
                        worksheet.SheetView.FreezeRows(1);
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
        #endregion

    }
}