using Helpers;
using Newtonsoft.Json.Linq;
using PDMDA;
using SCMDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class TakeOrderController : WebController
    {
        private ProductDA productDA = new ProductDA();
        private TakeOrderDA takeOrderDA = new TakeOrderDA();

        public ActionResult TakeOrder()
        {
            return View();
        }

        #region//Get
        #endregion

        #region//Add
        #region//AddDemand --需求單單頭 新增 -- Ding 2022.12.26
        [HttpPost]
        public void AddDemand(string DemandNo = "", string CompanyNo = "", string DemandSource = "", string DemandDesc = "", string CustNo = "",
            string DemandStatus = "", string Remark = "", string DemandDate = "")
        {
            try
            {
                #region //Request
                dataRequest = takeOrderDA.AddDemand(DemandNo, CompanyNo, DemandSource, DemandDesc, CustNo, Remark, DemandDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDemandLine --需求單單身 新增 -- Ding 2022.12.26
        [HttpPost]
        public void AddDemandLine(string DemandLineNo = "", int DemandId = -1, int ProdTypeId = -1, string TypeNo = "",
            string WipCategory = "", string CustMtlName = "", string DeliveryDate = "", int UnitPrice = -1,
            string UomNo = "", int OrderQty = -1, string OrderType = "", string MtlSpec = "", string DeliveryProcess = "", string Remark = "")
        {
            try
            {
                #region //Request
                dataRequest = takeOrderDA.AddDemandLine(DemandLineNo, DemandId, ProdTypeId, TypeNo,
                    WipCategory, CustMtlName, DeliveryDate, UnitPrice,
                    UomNo, OrderQty, OrderType, MtlSpec, DeliveryProcess, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//Update
        #endregion

        #region//Delete
        #endregion

        #region//Api

        #region//訂單
        #endregion

        #region//品號
        [HttpPost]
        [Route("api/DemandIssuance/MtlItem")]
        public void AddMtlItem(int DemandId = -1, int DemandLineId = -1, int MtlItemId = -1, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", int InventoryUomId = -1, int WeightUomId = -1, int PurchaseUomId = -1, int SaleUomId = -1
            , string TypeOne = "", string TypeTwo = "", string TypeThree = "", string TypeFour = "", int InventoryId = -1, int RequisitionInventoryId = -1, string InventoryManagement = "", string MtlModify = ""
            , string BondedStore = "", string ItemAttribute = "", string LotManagement = "", string MeasureType = "", string OverReceiptManagement = "", string OverDeliveryManagement = "", string EffectiveDate = ""
            , string ExpirationDate = "", string MtlItemDesc = "", string MtlItemRemark = "", string CustomerMtlItems = "")
        {
            try
            {
                #region //Request
                //dataRequest = productDA.AddMtlItem(MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, InventoryUomId, WeightUomId, PurchaseUomId, SaleUomId
                //    , TypeOne, TypeTwo, TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryManagement, MtlModify, BondedStore, ItemAttribute
                //    , LotManagement, MeasureType, OverReceiptManagement, OverDeliveryManagement, EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark, CustomerMtlItems);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//Bom
        #endregion

        #region//圖面管理
        #endregion

        #region//工單
        //5902的工單不能開
        #endregion

        #region//途程        
        #endregion

        #region//製令        
        #endregion

        #endregion
    }
}