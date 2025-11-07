using Helpers;
using Newtonsoft.Json.Linq;
using PDMDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class RdWorkManagementController : WebController
    {
        private RdWorkBasicInformationDA rdWorkBasicInformationDA = new RdWorkBasicInformationDA();

        #region //View
        public ActionResult RdWorkPlatform()
        {
            ViewLoginCheck();

            return View();
        }

        [HttpPost]
        public ActionResult EditMtlItem(int MtlItemId) {
            //ViewBag.MtlItemId = MtlItemId;
            return PartialView("~/Views/RdWorkManagement/EditMtlItem.cshtml");
        }

        #endregion

        #region //Get
        #region //GetBillOfMaterialsTree 取得Bom結構資料
        [HttpPost]
        public void GetBillOfMaterialsTree(int BomId = -1, int MtlItemId = -1, string MtlItemNo = "", string MtlItemName = "", string MtlItemSpec = "", string CustomerMtlItemNo = ""
            , string Status = "", string TransferStatus = "", string BomTransferStatus = "", string CustomerTransferStatus = "", string BomSubstitutionTransferStatus = "", string UomCTransferStatus = ""
            , string CheckMtlDate = "", int ItemTypeId = -1, string LotManagement = "", string ItemAttribute = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "read,constrained-data");
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                #region //Request
                dataRequest = rdWorkBasicInformationDA.GetBillOfMaterialsTree(BomId, MtlItemId, MtlItemNo, MtlItemName, MtlItemSpec, CustomerMtlItemNo
                    , Status, TransferStatus, BomTransferStatus, CustomerTransferStatus, BomSubstitutionTransferStatus, UomCTransferStatus, CheckMtlDate
                    , ItemTypeId, LotManagement, ItemAttribute);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.GetBillOfMaterials(BomId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void GetBomDetail(int BomDetailId = -1, int BomId = -1, int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.GetBomDetail(BomDetailId, BomId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CheckExistsMtlItemNo 判斷品號是否存在
        [HttpPost]
        public void CheckExistsMtlItemNo(string BomStructureJson = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "add");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.CheckExistsMtlItemNo(BomStructureJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddBomMtlItemStruct 品號BOM結構新增
        [HttpPost]
        public void AddBomStructure(string BomStructureJson = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "add");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.AddBomStructure(BomStructureJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMtlItemBatchExcel 品號批量新增
        [HttpPost]
        public void AddMtlItemBatchExcel(string SpreadsheetData = "", string ItemTypeSegmentList = "", bool CustomizeMtlItemNo = false)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "add");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.AddMtlItemBatchExcel(SpreadsheetData, ItemTypeSegmentList, CustomizeMtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
            , double LossRate = 0.0, string StandardCostingType = "", string MaterialProperties = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.AddBomDetail(BomId, MtlItemId, BomSequence, CompositionQuantity, Base, LossRate, StandardCostingType, MaterialProperties, EffectiveDate, ExpirationDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //TransferToERP 品號BOM結構批量拋轉ERP
        [HttpPost]
        public void TransferToERP(int MtlItemId)
        {
            try
            {
                WebApiLoginCheck("MtlItem", "data-transfer");

                #region //Request
                dataRequest = rdWorkBasicInformationDA.TransferToERP(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateBillOfMaterials Bom主件資料更新
        [HttpPost]
        public void UpdateBillOfMaterials(int BomId = -1, int MtlItemId = -1, double StandardLot = 0.0, string WipPrefix = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.UpdateBillOfMaterials(BomId, MtlItemId, StandardLot, WipPrefix, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
            , string StandardCostingType = "", string MaterialProperties = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("MtlItem", "bom");

                #region //Request
                rdWorkBasicInformationDA = new RdWorkBasicInformationDA();
                dataRequest = rdWorkBasicInformationDA.UpdateBomDetail(BomDetailId, BomId, MtlItemId, BomSequence, CompositionQuantity, Base, LossRate
                    , StandardCostingType, MaterialProperties, EffectiveDate, ExpirationDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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