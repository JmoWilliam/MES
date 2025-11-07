using ClosedXML.Excel;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PDMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class PdmBasicInformationController : WebController
    {
        private PdmBasicInformationDA pdmBasicInformationDA = new PdmBasicInformationDA();

        #region //View
        public ActionResult MtlElement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UnitOfMeasure()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UomCalculate()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MtlModel()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MtlPrinciple()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ItemSegmentManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ItemTypeManagment()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult DfmParameterManagment()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult DfmTemplate()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult DfmTemplateParameter()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ViewDfmItem()
        {
            return View();
        }
        public ActionResult DfmItemCategoryManagment()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult DfmMaterialManagment()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult DfmCategoryMaterial()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult ViewMaterialManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MoldingConditions()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ItemCategoryDept()
        {
            ViewLoginCheck();
            return View();
        }
         
        #endregion

        #region //Get
        #region //GetMtlElement 取得品號元素資料
        [HttpPost]
        public void GetMtlElement(int ElementId = -1, string ElementNo = "", string ElementName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlElement", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetMtlElement(ElementId, ElementNo, ElementName, Status
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

        #region //GetUnitOfMeasure 取得單位基本資料
        [HttpPost]
        public void GetUnitOfMeasure(int UomId = -1, string UomNo = "", string UomName = "", string UomType = ""
            , string Status = "", int RejectUomId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetUnitOfMeasure(UomId, UomNo, UomName, UomType
                    , Status, RejectUomId
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

        #region //GetUomCalculate 取得單位換算資料
        [HttpPost]
        public void GetUomCalculate(int UomCalculateId = -1, string ConvertUomNo = "", string ConvertUomName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "calculate");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetUomCalculate(UomCalculateId, ConvertUomNo, ConvertUomName, Status
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

        #region //GetMtlModel 取得品號機型資料
        [HttpPost]
        public void GetMtlModel(int MtlModelId = -1, int ParentId = -2, string MtlModelNo = "", string MtlModelName = "", string Status = ""
            , string QueryType = "Query", int StartParent = -1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MtlModel", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetMtlModel(MtlModelId, ParentId, MtlModelNo, MtlModelName, Status
                    , QueryType, StartParent, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrinciple 取得品號原則
        [HttpPost]
        public void GetPrinciple(int PrincipleId = -1, int BuildId = -1, string Value = "", string OrderBy = "")
        {
            try
            {
                WebApiLoginCheck("MtlModel", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetPrinciple(PrincipleId, BuildId, Value, OrderBy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrincipleCondition 取得品號原則條件
        [HttpPost]
        public void GetPrincipleCondition(int ConditionId = -1, int BuildId = -1, string PrincipleValue = "")
        {
            try
            {
                WebApiLoginCheck("MtlModel", "read");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetPrincipleCondition(ConditionId, BuildId, PrincipleValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetItemSegment 取得編碼節段資料
        [HttpPost]
        public void GetItemSegment(int ItemSegmentId = -1, string SegmentNo = "", string SegmentName = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemSegment(ItemSegmentId, SegmentNo, SegmentName, StartDate, EndDate, Status
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

        #region //GetItemSegmentSimple 取得編碼節段資料(無換頁)
        [HttpPost]
        public void GetItemSegmentSimple(string Today = "")
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemSegmentSimple(Today);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetItemSegmentValue 取得編碼節段資料
        [HttpPost]
        public void GetItemSegmentValue(int SegmentValueId = -1, int ItemSegmentId = -1, string SegmentValueNo = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "detail,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemSegmentValue(SegmentValueId, ItemSegmentId, SegmentValueNo, StartDate, EndDate, Status
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

        #region //GetItemSegmentValueSimple 取得編碼節段Value資料(無換頁)
        [HttpPost]
        public void GetItemSegmentValueSimple(int ItemSegmentId = -1, string SegmentValueNo = "", string Today = "")
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemSegmentValueSimple(ItemSegmentId, SegmentValueNo, Today);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetItemType 取得品號類別資料
        [HttpPost]
        public void GetItemType(int ItemTypeId = -1, string ItemTypeNo = "", string ItemTypeName = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemType(ItemTypeId, ItemTypeNo, ItemTypeName, StartDate, EndDate, Status
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

        #region //GetItemTypeSimple 取得品號類別資料(無換頁)
        [HttpPost]
        public void GetItemTypeSimple(string Today = "")
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemTypeSimple(Today);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetItemTypeSegment 取得品號類別結構資料
        [HttpPost]
        public void GetItemTypeSegment(int ItemTypeSegmentId = -1, int ItemTypeId = -1, string SegmentType = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "detail,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemTypeSegment(ItemTypeSegmentId, ItemTypeId, SegmentType, StartDate, EndDate, Status
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

        #region //GetItemTypeDefault 取得品號類別預設資料
        [HttpPost]
        public void GetItemTypeDefault(int ItemTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "detail,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemTypeDefault(ItemTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmItem 取得DFM 參數
        [HttpPost]
        public void GetDfmItem(int DfmItemId=-1, string DfmItemNo="", string DfmItemName="",string Status=""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DfmParameterManagment", "read,constrained-data");

                #region //Request
                //dataRequest = pdmBasicInformationDA.GetDfmItem(DfmItemId, DfmItemNo, DfmItemName, Status
                //    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmTemplate 取的DFM樣板資訊
        [HttpPost]
        public void GetDfmTemplate(int DfmTemplateId=-1, string DfmTempNo = "", string DfmTempName = "", int RfqProTypeId=-1,int ProductUseId=-1,string Status=""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DfmParameterManagment", "read,constrained-data");

                #region //Request
                //dataRequest = pdmBasicInformationDA.GetDfmTemplate(DfmTemplateId,DfmTempNo, DfmTempName, RfqProTypeId, ProductUseId, Status
                //    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmItemCategory --取得DFM項目種類代碼
        [HttpPost]
        public void GetDfmItemCategory(int DfmItemCategoryId = -1, int ModeId = -1, string DfmItemCategoryNo = "", string DfmItemCategoryName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "read");

                #region //Request 
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.GetDfmItemCategory(DfmItemCategoryId, ModeId, DfmItemCategoryNo, DfmItemCategoryName, Status
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

        #region//GetDfmCategoryTemplate --取得DFM項目種類樣板
        [HttpPost]
        public void GetDfmCategoryTemplate(int DfmCtId = -1, int DfmItemCategoryId = -1, string DataType = "")
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "read");

                #region //Request 
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.GetDfmCategoryTemplate(DfmCtId, DfmItemCategoryId, DataType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmMaterial --取得DFM物料
        [HttpPost]
        public void GetDfmMaterial(int DfmMaterialId = -1, string DfmMaterialNo = "", string DfmMaterialName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "read");

                #region //Request 
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.GetDfmMaterial(DfmMaterialId, DfmMaterialNo, DfmMaterialName, Status
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

        #region//GetDfmCategoryMaterial --取得DFM項目種類物料
        [HttpPost]
        public void GetDfmCategoryMaterial(string DfmMaterialDesc="", int DfmMaterialId = -1, int DfmItemCategoryId = -1, string DfmMaterialNo = "", string DfmMaterialName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "read");

                #region //Request 
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.GetDfmCategoryMaterial(DfmMaterialDesc, DfmMaterialId, DfmItemCategoryId, DfmMaterialNo, DfmMaterialName, Status
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

        #region //GetMoldingConditions 取得成型產品生產條件
        [HttpPost]
        public void GetMoldingConditions(int McId = -1, string MtlItemNo = "", string MtlItemName = "", string Material = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetMoldingConditions(McId, MtlItemNo, MtlItemName, Material
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

        #region //GetProductionStockInDetail 取得ERP生產入庫明細表
        [HttpPost]
        public void GetProductionStockInDetail(string MtlItemNo = "", string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetProductionStockInDetail(MtlItemNo,StardDay, EndDay);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTemporaryOrderOutInDetail ERP暫出入單明細表
        [HttpPost]
        public void GetTemporaryOrderOutInDetail(string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetTemporaryOrderOutInDetail(StardDay, EndDay);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTempOutReturnRecordDetail ERP暫出歸還單明細表
        [HttpPost]
        public void GetTempOutReturnRecordDetail(string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetTempOutReturnRecordDetail(StardDay, EndDay);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetItemCategoryDept 取得品號類別責任部門
        [HttpPost]
        public void GetItemCategoryDept(int ItemCatDeptId = -1, string TypeTwo = "", string TypeThree = "", string InventoryNo = "", string CatDept = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ItemCategoryDept", "read,constrained-data");

                #region //Request
                dataRequest = pdmBasicInformationDA.GetItemCategoryDept(ItemCatDeptId,TypeTwo, TypeThree, InventoryNo, CatDept
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
        #region //AddMtlElement 品號元素資料新增
        [HttpPost]
        public void AddMtlElement(string ElementNo = "", string ElementName = "")
        {
            try
            {
                WebApiLoginCheck("MtlElement", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddMtlElement(ElementNo, ElementName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddUnitOfMeasure 單位基本資料新增
        [HttpPost]
        public void AddUnitOfMeasure(string UomNo = "", string UomName = "", string UomType = "", string UomDesc = "")
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddUnitOfMeasure(UomNo, UomName, UomType, UomDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddUomCalculate 單位換算資料新增
        [HttpPost]
        public void AddUomCalculate(int UomId = -1, int ConvertUomId = -1, double Value = 0)
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "calculate");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddUomCalculate(UomId, ConvertUomId, Value);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMtlModel 品號機型資料新增
        [HttpPost]
        public void AddMtlModel(int ParentId = -1, string MtlModelNo = ""
            , string MtlModelName = "", string MtlModelDesc = "")
        {
            try
            {
                WebApiLoginCheck("MtlModel", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddMtlModel(ParentId, MtlModelNo
                    , MtlModelName, MtlModelDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddItemSegment 編碼節段新增
        [HttpPost]
        public void AddItemSegment(string SegmentNo = "", string SegmentName = "", string SegmentDesc = "", string EffectiveDate = "", string InactiveDate = ""
            )
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddItemSegment(SegmentNo, SegmentName, SegmentDesc, EffectiveDate, InactiveDate
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

        #region //AddSegmentValue 編碼節段Value新增
        [HttpPost]
        public void AddSegmentValue(int ItemSegmentId = -1, string SegmentValueNo = "", string SegmentValue = "", string SegmentValueDesc = "", string EffectiveDate = "", string InactiveDate = ""
            )
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "detail");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddSegmentValue(ItemSegmentId, SegmentValueNo, SegmentValue, SegmentValueDesc, EffectiveDate, InactiveDate
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

        #region //AddItemType 品號類別新增
        [HttpPost]
        public void AddItemType(string ItemTypeNo = "", string ItemTypeName = "", string ItemTypeDesc = "", string EffectiveDate = "", string InactiveDate = ""
            )
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddItemType(ItemTypeNo, ItemTypeName, ItemTypeDesc, EffectiveDate, InactiveDate
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

        #region //AddItemTypeSegment 品號類別結構新增
        [HttpPost]
        public void AddItemTypeSegment(int ItemTypeId = -1, string SegmentType = "", string SegmentValue = "", string SuffixCode = "", string EffectiveDate = "", string InactiveDate = ""
            )
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddItemTypeSegment(ItemTypeId, SegmentType, SegmentValue, SuffixCode, EffectiveDate, InactiveDate
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

        #region //AddItemTypeDefault 品號類別預設資料新增
        [HttpPost]
        public void AddItemTypeDefault(int ItemTypeId = -1, string TypeOne = "", string TypeTwo = "", string TypeThree = "", string TypeFour = ""
            , int InventoryId = -1, int RequisitionInventoryId = -1, int InventoryUomId = -1, string ItemAttribute = ""
            , string MeasureType = "", int PurchaseUomId = -1, int SaleUomId = -1, string LotManagement = "", string InventoryManagement = ""
            , string MtlModify = "", string BondedStore = "", string OverReceiptManagement = "", string OverDeliveryManagement = ""
            , string EffectiveDate = "", string ExpirationDate = "", string MtlItemDesc = "", string MtlItemRemark = ""
            )
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "add");

                #region //Request
                dataRequest = pdmBasicInformationDA.AddItemTypeDefault(ItemTypeId, TypeOne, TypeTwo, TypeThree, TypeFour
                    , InventoryId, RequisitionInventoryId, InventoryUomId, ItemAttribute
                    , MeasureType, PurchaseUomId, SaleUomId, LotManagement, InventoryManagement
                    , MtlModify, BondedStore, OverReceiptManagement, OverDeliveryManagement
                    , EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark
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

        #region//AddDfmItem 新增DFM參數
        [HttpPost]
        public void AddDfmItem(string DfmItemNo = "", string DfmItemName = "")
        {
            try
            {
                WebApiLoginCheck("DfmParameterManagment", "add");

                #region //Request
                //dataRequest = pdmBasicInformationDA.AddDfmItem(DfmItemNo, DfmItemName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmTemplate 新增DFM樣板 單頭
        [HttpPost]
        public void AddDfmTemplate(string DfmTempParamName = "", int RfqProTypeId = -1,int ProductUseId=-1)
        {
            try
            {
                WebApiLoginCheck("DfmTemplate", "add");

                #region //Request
                //dataRequest = pdmBasicInformationDA.AddDfmTemplate(DfmTempParamName, RfqProTypeId, ProductUseId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region// AddDfmTemplateParameter 新增DFM樣板 單身
        #endregion

        #region//AddDfmItemCategory --建立DFM種類代碼
        [HttpPost]
        public void AddDfmItemCategory(int ModeId = -1, string DfmItemCategoryNo = "", string DfmItemCategoryName = "", string DfmItemCategoryDesc = "")
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "add");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.AddDfmItemCategory(ModeId, DfmItemCategoryNo, DfmItemCategoryName, DfmItemCategoryDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmMaterial --建立DFM物料代碼
        [HttpPost]
        public void AddDfmMaterial(string DfmMaterialNo = "", string DfmMaterialName = "", string DfmMaterialDesc = "", double DfmMaterialMoney = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "add");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.AddDfmMaterial(DfmMaterialNo, DfmMaterialName, DfmMaterialDesc, DfmMaterialMoney);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmCategoryMaterial --建立DFM項目種類物料
        [HttpPost]
        public void AddDfmCategoryMaterial(int DfmMaterialId=-1, string ProcessData="", int DfmItemCategoryId = -1, string DfmMaterialNo = "", string DfmMaterialName = "", string DfmMaterialDesc = "", double DfmMaterialMoney = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "add");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.AddDfmCategoryMaterial(DfmMaterialId, ProcessData, DfmItemCategoryId, DfmMaterialNo, DfmMaterialName, DfmMaterialDesc, DfmMaterialMoney);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ImportExcelMoldingConditions 解析Excel新增成形產品條件
        [HttpPost]
        public void ImportExcelMoldingConditions(string FileId = "")
        {
            try
            {
                if (FileId.Length <= 0) throw new SystemException("必須上傳檔案!!");

                List<StockInDetail> mcExcelFormats = new List<StockInDetail>();

                string[] FileIds = FileId.Split(',');
                foreach (var fileId in FileIds)
                {
                    #region //取得File資料
                    BmFileInfo fileInfo = pdmBasicInformationDA.GetFileInfoById(Convert.ToInt32(fileId));
                    if (fileInfo == null)
                    {
                        throw new SystemException("查無此File資料!!");
                    }
                    #endregion

                    #region //解析EXCEL
                    using (var stream = new MemoryStream(fileInfo.FileContent))
                    {
                        using (XLWorkbook workbook = new XLWorkbook(stream))
                        {
                            IXLWorksheet worksheet = workbook.Worksheet(1);
                            var firstCell = worksheet.FirstCellUsed();
                            var lastCell = worksheet.LastCellUsed();

                            int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                            var data = worksheet.Range(firstCell.Address, lastCell.Address);
                            var table = data.AsTable();

                            for (var i = 2; i <= CellLength; i++)
                            {
                                try
                                {
                                    StockInDetail mcExcelFormat = new StockInDetail()
                                    {
                                        MtlItemNo = table.Cell(i, 1).Value.ToString(),
                                        Material = table.Cell(i, 2).Value.ToString(),
                                        CycleTime = table.Cell(i, 3).Value.ToString(),
                                        MoldWeight = table.Cell(i, 4).Value.ToString(),
                                        Cavity = table.Cell(i, 5).Value.ToString(),
                                        UnitPrice = Convert.ToDouble(table.Cell(i, 6).Value),
                                        ProcessingFee  = Convert.ToDouble(table.Cell(i, 7).Value),
                                        TheoreticalQty = Convert.ToDouble(table.Cell(i, 8).Value),
                                    };
                                    mcExcelFormats.Add(mcExcelFormat);
                                }
                                catch (Exception ex)
                                {
                                    throw new SystemException("第" + i + "筆資料ERROR:" + ex.Message);
                                }
                            }
                        }
                    }
                    #endregion
                }

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.AddBatchMoldingConditions(mcExcelFormats);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    jsonResponse = dataRequestJson;
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

        #region //AddItemCategoryDept 新增品號類別責任部門 
        [HttpPost]
        public void AddItemCategoryDept(string TypeTwo = "", string TypeThree = "", string InventoryNo = "", string CatDept = "")
        {
            try
            {
                WebApiLoginCheck("ItemCategoryDept", "add");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.AddItemCategoryDept(TypeTwo, TypeThree, InventoryNo, CatDept);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMoldingConditions 新增成型產品生產條件 
        [HttpPost]
        public void AddMoldingConditions( string MtlItemId = "",  string Material = "", string CycleTime = "", string MoldWeight = ""
            , string Cavity = "", string ProcessingFee = "", string TheoreticalQty = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "add");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.AddMoldingConditions(MtlItemId, Material, CycleTime, MoldWeight, Cavity, ProcessingFee, TheoreticalQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateMtlElement 品號元素資料更新
        [HttpPost]
        public void UpdateMtlElement(int ElementId = -1, string ElementNo = "", string ElementName = "")
        {
            try
            {
                WebApiLoginCheck("MtlElement", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateMtlElement(ElementId, ElementNo, ElementName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlElementStatus 品號元素狀態更新
        [HttpPost]
        public void UpdateMtlElementStatus(int ElementId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlElement", "status-switch");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateMtlElementStatus(ElementId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlElementSort 品號元素順序調整
        [HttpPost]
        public void UpdateMtlElementSort(string MtlElementList = "")
        {
            try
            {
                WebApiLoginCheck("MtlElement", "sort");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateMtlElementSort(MtlElementList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUnitOfMeasure 單位基本資料更新
        [HttpPost]
        public void UpdateUnitOfMeasure(int UomId = -1, string UomNo = "", string UomName = ""
            , string UomType = "", string UomDesc = "")
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateUnitOfMeasure(UomId, UomNo, UomName
                    , UomType, UomDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUnitOfMeasureStatus 單位基本資料狀態更新
        [HttpPost]
        public void UpdateUnitOfMeasureStatus(int UomId = -1)
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "status-switch");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateUnitOfMeasureStatus(UomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUomCalculate 單位換算資料更新
        [HttpPost]
        public void UpdateUomCalculate(int UomCalculateId = -1, int UomId = -1, int ConvertUomId = -1, double Value = 0)
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "calculate");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateUomCalculate(UomCalculateId, UomId, ConvertUomId, Value);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUomCalculateStatus 單位換算資料狀態更新
        [HttpPost]
        public void UpdateUomCalculateStatus(int UomCalculateId = -1)
        {
            try
            {
                WebApiLoginCheck("UnitOfMeasure", "calculate");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateUomCalculateStatus(UomCalculateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlModel 品號機型資料更新
        [HttpPost]
        public void UpdateMtlModel(int MtlModelId = -1, string MtlModelNo = "", string MtlModelName = "", string MtlModelDesc = "")
        {
            try
            {
                WebApiLoginCheck("MtlModel", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateMtlModel(MtlModelId, MtlModelNo, MtlModelName, MtlModelDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlModelStatus 品號機型狀態更新
        [HttpPost]
        public void UpdateMtlModelStatus(int MtlModelId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlModel", "status-switch");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateMtlModelStatus(MtlModelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMtlModelSort 品號機型順序調整
        [HttpPost]
        public void UpdateMtlModelSort(int ParentId = -1, string MtlModelList = "")
        {
            try
            {
                WebApiLoginCheck("MtlModel", "sort");
            
                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateMtlModelSort(ParentId, MtlModelList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemSegment 編碼節段更新
        [HttpPost]
        public void UpdateItemSegment(int ItemSegmentId = -1, string SegmentName = "", string SegmentDesc = "", string EffectiveDate = "", string InactiveDate = "")
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemSegment(ItemSegmentId, SegmentName, SegmentDesc, EffectiveDate, InactiveDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemSegmentIdStatus 編碼節段狀態更新
        [HttpPost]
        public void UpdateItemSegmentIdStatus(int ItemSegmentId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemSegmentIdStatus(ItemSegmentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSegmentValue 編碼節段Value更新
        [HttpPost]
        public void UpdateSegmentValue(int ItemSegmentId = -1, int SegmentValueId = -1, string SegmentValue = "", string SegmentValueDesc = "", string EffectiveDate = "", string InactiveDate = "")
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "detail");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateSegmentValue(ItemSegmentId, SegmentValueId, SegmentValue, SegmentValueDesc, EffectiveDate, InactiveDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemSegmentValueStatus 編碼節段狀態更新
        [HttpPost]
        public void UpdateItemSegmentValueStatus(int ItemSegmentId,int SegmentvalueId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemSegmentManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemSegmentValueStatus(ItemSegmentId,SegmentvalueId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemType 品號類別更新
        [HttpPost]
        public void UpdateItemType(int ItemTypeId = -1, string ItemTypeNo = "", string ItemTypeName = "", string ItemTypeDesc = "", string EffectiveDate = "", string InactiveDate = "")
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemType(ItemTypeId, ItemTypeNo, ItemTypeName, ItemTypeDesc, EffectiveDate, InactiveDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemTypeIdStatus 品號類別狀態更新
        [HttpPost]
        public void UpdateItemTypeIdStatus(int ItemTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemTypeIdStatus(ItemTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemTypeSegment 品號類別結構新增更新
        [HttpPost]
        public void UpdateItemTypeSegment(int ItemTypeId = -1, int ItemTypeSegmentId = -1, string SegmentType = "", string SegmentValue = "", string SuffixCode = "", string EffectiveDate = "", string InactiveDate = "")
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemTypeSegment(ItemTypeId, ItemTypeSegmentId, SegmentType, SegmentValue, SuffixCode, EffectiveDate, InactiveDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemTypeSegmentSort 品號類別結構新增更新
        [HttpPost]
        public void UpdateItemTypeSegmentSort(string ItemTypeSegmentIdList = "")
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemTypeSegmentSort(ItemTypeSegmentIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemTypeSegmentIdStatus 品號類別結構狀態更新
        [HttpPost]
        public void UpdateItemTypeSegmentIdStatus(int ItemTypeId = -1, int ItemTypeSegmentId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemTypeSegmentIdStatus(ItemTypeId, ItemTypeSegmentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemTypeDefault 品號類別預設資料更新
        [HttpPost]
        public void UpdateItemTypeDefault(int ItemTypeDefaultId = -1, int ItemTypeId = -1, string TypeOne = "", string TypeTwo = "", string TypeThree = "", string TypeFour = ""
            , int InventoryId = -1, int RequisitionInventoryId = -1, int InventoryUomId = -1, string ItemAttribute = ""
            , string MeasureType = "", int PurchaseUomId = -1, int SaleUomId = -1, string LotManagement = "", string InventoryManagement = ""
            , string MtlModify = "", string BondedStore = "", string OverReceiptManagement = "", string OverDeliveryManagement = ""
            , string EffectiveDate = "", string ExpirationDate = "", string MtlItemDesc = "", string MtlItemRemark = ""
            )
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "update");

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateItemTypeDefault(ItemTypeDefaultId, ItemTypeId, TypeOne, TypeTwo, TypeThree, TypeFour
                    , InventoryId, RequisitionInventoryId, InventoryUomId, ItemAttribute
                    , MeasureType, PurchaseUomId, SaleUomId, LotManagement, InventoryManagement
                    , MtlModify, BondedStore, OverReceiptManagement, OverDeliveryManagement
                    , EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark
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

        #region//UpdateDfmItem DFM參數更新
        [HttpPost]
        public void UpdateDfmItem(int DfmItemId = -1, string DfmItemNo = "", string DfmItemName = "")
        {
            try
            {
                WebApiLoginCheck("DfmParameterManagment", "update");

                #region //Request
                //dataRequest = pdmBasicInformationDA.UpdateDfmItem(DfmItemId, DfmItemNo, DfmItemName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmItemStatus DFM參數狀態更新
        [HttpPost]
        public void UpdateDfmItemStatus(int DfmItemId = -1)
        {
            try
            {
                WebApiLoginCheck("DfmParameterManagment", "status-switch");

                #region //Request
                //dataRequest = pdmBasicInformationDA.UpdateDfmItemStatus(DfmItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmTemplate DFM樣板更新
        [HttpPost]
        public void UpdateDfmTemplate(int DfmTemplateId=-1, string DfmTempParamName = "", int RfqProTypeId = -1, int ProductUseId = -1)
        {
            try
            {
                WebApiLoginCheck("DfmTemplate", "update");

                #region //Request
                //dataRequest = pdmBasicInformationDA.UpdateDfmTemplate(DfmTemplateId,  DfmTempParamName, RfqProTypeId, ProductUseId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmTemplateStatus DFM樣板狀態更新
        [HttpPost]
        public void UpdateDfmTemplateStatus(int DfmTemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("DfmTemplate", "status-switch");

                #region //Request
                //dataRequest = pdmBasicInformationDA.UpdateDfmTemplateStatus(DfmTemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmItemCategory --更新DFM項目種類代碼
        [HttpPost]
        public void UpdateDfmItemCategory(int DfmItemCategoryId = -1, int ModeId = -1, string DfmItemCategoryName = "", string DfmItemCategoryDesc = "")
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "update");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateDfmItemCategory(DfmItemCategoryId, ModeId, DfmItemCategoryName, DfmItemCategoryDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmItemCategoryStatus --更新DFM項目種類代碼狀態
        [HttpPost]
        public void UpdateDfmItemCategoryStatus(int DfmItemCategoryId)
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "status-switch");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateDfmItemCategoryStatus(DfmItemCategoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmCategoryTemplate -更新DfmCategoryTemplate資料
        [HttpPost]
        public void UpdateDfmCategoryTemplate(int DfmCtId =-1, int DfmItemCategoryId = -1, string DataType = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "update");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateDfmCategoryTemplate1(DfmCtId,DfmItemCategoryId, DataType, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmMaterial --更新DFM物料代碼
        [HttpPost]
        public void UpdateDfmMaterial(int DfmMaterialId = -1, string DfmMaterialName = "", string DfmMaterialDesc = "", double DfmMaterialMoney = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "update");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateDfmMaterial(DfmMaterialId, DfmMaterialName, DfmMaterialDesc, DfmMaterialMoney);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmMaterialStatus --更新DFM物料代碼狀態
        [HttpPost]
        public void UpdateDfmMaterialStatus(int DfmMaterialId)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "status-switch");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateDfmMaterialStatus(DfmMaterialId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmCategoryMaterialStatus --更新DFM項目種類物料狀態
        [HttpPost]
        public void UpdateDfmCategoryMaterialStatus(int DfmCategoryMaterialId)
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "category-material");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateDfmCategoryMaterialStatus(DfmCategoryMaterialId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateMoldingConditionsUnitPrice --更新成型條件品號採購單價
        [HttpPost]
        public void UpdateMoldingConditionsUnitPrice()
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "purchase-price");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateMoldingConditionsUnitPrice();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemCategoryDept 更新品號類別責任部門
        [HttpPost]
        public void UpdateItemCategoryDept(int ItemCatDeptId = -1, string TypeTwo = "", string TypeThree = "", string InventoryNo = "", string CatDept = "")
        {
            try
            {
                WebApiLoginCheck("ItemCategoryDept", "update");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateItemCategoryDept(ItemCatDeptId, TypeTwo, TypeThree, InventoryNo, CatDept);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMoldingConditions 更新成型產品生產條件 
        [HttpPost]
        public void UpdateMoldingConditions(int McId = -1, string MtlItemId = "", string Material = "", string CycleTime = "", string MoldWeight = ""
            , string Cavity = "", string ProcessingFee = "", string TheoreticalQty = "")
        {
            try
            {
                WebApiLoginCheck("ItemCategoryDept", "update");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.UpdateMoldingConditions(McId, MtlItemId, Material, CycleTime, MoldWeight, Cavity, ProcessingFee, TheoreticalQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteMtlModel -- 品號機型資料刪除
        [HttpPost]
        public void DeleteMtlModel(int MtlModelId = -1)
        {
            try
            {
                WebApiLoginCheck("MtlModel", "delete");

                #region //Request
                dataRequest = pdmBasicInformationDA.DeleteMtlModel(MtlModelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteItemTypeDefault -- 品號類別預設資料刪除
        [HttpPost]
        public void DeleteItemTypeDefault(int ItemTypeDefaultId = -1, int ItemTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemTypeManagment", "delete");

                #region //Request
                dataRequest = pdmBasicInformationDA.DeleteItemTypeDefault(ItemTypeDefaultId, ItemTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmItemCategory --刪除DFM項目種類代碼
        [HttpPost]
        public void DeleteDfmItemCategory(int DfmItemCategoryId = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "delete");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.DeleteDfmItemCategory(DfmItemCategoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmCategoryTemplate --刪除DFM項目種類樣板
        [HttpPost]
        public void DeleteDfmCategoryTemplate(int DfmItemCategoryId = -1, string DataType = "")
        {
            try
            {
                WebApiLoginCheck("DfmItemCategoryManagment", "delete");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.DeleteDfmCategoryTemplate(DfmItemCategoryId, DataType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmMaterial --刪除DFM物料代碼
        [HttpPost]
        public void DeleteDfmMaterial(int DfmQiMaterialId = -1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "delete");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.DeleteDfmMaterial(DfmQiMaterialId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmCategoryMaterial --刪除DFM項目種類物料
        [HttpPost]
        public void DeleteDfmCategoryMaterial(int DfmCategoryMaterialId=-1)
        {
            try
            {
                WebApiLoginCheck("DfmMaterialManagment", "delete");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.DeleteDfmCategoryMaterial(DfmCategoryMaterialId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteItemCatDept 刪除品號類別責任部門
        [HttpPost]
        public void DeleteItemCatDept(int ItemCatDeptId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemCategoryDept", "delete");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.DeleteItemCatDept(ItemCatDeptId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMoldingConditions 刪除成型產品生產條件
        [HttpPost]
        public void DeleteMoldingConditions(int McId = -1)
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "delete");

                #region //Request
                pdmBasicInformationDA = new PdmBasicInformationDA();
                dataRequest = pdmBasicInformationDA.DeleteMoldingConditions(McId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateUnitOfMeasureSynchronize -- 單位資料同步
        [HttpPost]
        [Route("api/ERP/UnitOfMeasureSynchronize")]
        public void UpdateUnitOfMeasureSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateUnitOfMeasureSynchronize");
                #endregion

                #region //Request
                dataRequest = pdmBasicInformationDA.UpdateUnitOfMeasureSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //ExcelProductionStockInDetail ERP生產入庫明細表輸出Excel
        public void ExcelProductionStockInDetail(string ViewMtlItemNo = "", string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                dataRequest = pdmBasicInformationDA.GetProductionStockInDetail(ViewMtlItemNo, StardDay, EndDay);
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
                    string excelFileName = "【MES2.0】成型定價周報表";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "入庫日", "製令", "品號", "品名", "規格", "週期", "模重", "穴數", "理論生產數"
                        , "入庫數", "收片率", "驗收數", "良品率", "報廢數", "報廢率"
                        , "理論材料成本", "單PCS材料成本", "良品材料成本", "不良品材料成本", "實際用料成本"
                        , "加工費", "理論產品單價", "人工製費", "實際營收", "成本合計", "毛利", "毛利率"
                        , "材料別稱", "材料品號", "材料品名", "材料規格", "採購單價" };
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
                            string ReceiptDate = item.ReceiptDate.ToString();
                            string ErpFull = item.ErpFull.ToString();
                            string MtlItemNo = item.MtlItemNo.ToString();
                            string MtlItemName = item.MtlItemName.ToString();
                            string MtlItemSpec = item.MtlItemSpec.ToString();
                            double CycleTime = Convert.ToDouble(item.CycleTime);
                            double MoldWeight = Convert.ToDouble(item.MoldWeight);
                            double Cavity = Convert.ToDouble(item.Cavity);
                            double TheoreticalQty = Convert.ToDouble(item.TheoreticalQty);
                            double ProcessingFee = Convert.ToDouble(item.ProcessingFee);
                            double ReceiptQty = Convert.ToDouble(item.ReceiptQty);
                            string YieldRate = (Math.Floor((ReceiptQty / TheoreticalQty) * 100)).ToString() + "%";
                            double AcceptQty = Convert.ToDouble(item.AcceptQty);
                            string PassRate = item.PassRate.ToString() + "%";
                            double ScriptQty = Convert.ToDouble(item.ScriptQty);
                            string DefectRate = item.DefectRate.ToString() + "%";
                            string Material = item.Material.ToString();
                            string CellMtlItemNo = item.CellMtlItemNo.ToString();
                            string CellMtlItemName = item.CellMtlItemName.ToString();
                            string CellMtlItemSpec = item.CellMtlItemSpec.ToString();
                            double UnitPrice = Convert.ToDouble(item.UnitPrice);
                            double ManCost = Convert.ToDouble(item.ManCost);
                            double MachineCost = Convert.ToDouble(item.MachineCost);
                            double LaborCost = Math.Floor((ManCost + MachineCost) * 1000) / 1000; //人工+制费 = 人時*人工單位成本+機時*機時單位成本

                            double theoPrice = Math.Floor((ProcessingFee / TheoreticalQty) * 1000) / 1000; ; //理論產品單價=加工費/理論生產數
                            double theoMatCost = Math.Floor((24 * 60 * 60 / CycleTime * MoldWeight / 1000 * UnitPrice) * 1000) / 1000; ; //理論材料成本 =24*60*60/週期*模重/1000*採購單價
                            double matCostPerPc = Math.Floor((theoMatCost / TheoreticalQty) * 1000) / 1000; ; //單PCS材料成本 = 理論材料成本/理論生產數
                            double actualRev = Math.Floor((theoPrice * AcceptQty) * 1000) / 1000; ; //實際營收 = 理論產品單價 * 验收数量
                            double goodMatCost = Math.Floor((AcceptQty * matCostPerPc) * 1000) / 1000; ; //良品材料成本 = 良品數 * 單PCS材料成本
                            double badMatCost = Math.Floor((ScriptQty * matCostPerPc) * 1000) / 1000; ; //不良品材料成本 = 不良品數 * 單PCS材料成本
                            double actualMatCost = Math.Floor((ReceiptQty * matCostPerPc) * 1000) / 1000; ; //實際用料成本 = 入庫數 * 單PCS材料成本
                            double grossProfit = Math.Floor((actualRev - LaborCost) * 1000) / 1000; ;//毛利 = 實際營收 - 人工+制费
                            string grossProfitRate = (Math.Floor((grossProfit / actualRev) * 1000) / 1000).ToString() + "%"; ;//毛利率 = 毛利/實際營收
                            double totalCost = Math.Floor((LaborCost + actualMatCost) * 1000) / 1000; ; //成本合計 = 人工+制费+實際用料成本

                            var resultList = new[]
                            {
                                new {
                                    ReceiptDate,
                                    ErpFull,
                                    MtlItemNo,
                                    MtlItemName,
                                    MtlItemSpec,
                                    CycleTime,
                                    MoldWeight,
                                    Cavity,
                                    TheoreticalQty,
                                    ReceiptQty,
                                    YieldRate,
                                    AcceptQty,
                                    PassRate,
                                    ScriptQty,
                                    DefectRate,
                                    theoMatCost,
                                    matCostPerPc,
                                    goodMatCost,
                                    badMatCost,
                                    actualMatCost,
                                    ProcessingFee,
                                    theoPrice,
                                    LaborCost,
                                    actualRev,
                                    totalCost,
                                    grossProfit,
                                    grossProfitRate,
                                    Material,
                                    CellMtlItemNo,
                                    CellMtlItemName,
                                    CellMtlItemSpec,
                                    UnitPrice,
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

        #region //ExcelTemporaryOrderOutInDetail ERP暫出入單明細表輸出Excel
        public void ExcelTemporaryOrderOutInDetail(string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                dataRequest = pdmBasicInformationDA.GetTemporaryOrderOutInDetail(StardDay, EndDay);
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
                    string excelFileName = "【MES2.0】暫出入單明細表";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "異動日期", "異動單號", "單據日期", "預計轉銷日", "對象代碼", "對象簡稱", "部門代號", "簡稱", "員工代號", "簡稱"
                        , "廠別代號", "簡稱", "幣別", "課稅別", "匯率", "件數", "備註", "品號", "品名", "規格" , "責任部門", "類別二", "類別三", "轉出庫別", "轉入庫別"
                        , "預計歸還日", "轉出儲位", "轉入儲位", "暫出數量", "暫出轉進銷量", "暫出歸還量", "暫入數量", "暫入轉進銷量", "暫入歸還量"
                        , "暫出包裝數量", "暫出包裝轉進銷量", "暫出包裝歸還量", "暫入包裝數量", "暫入包裝轉進銷量", "暫入包裝歸還量"
                        , "暫出贈/備品量", "暫出轉銷贈/備品量", "暫出歸還贈/備品量", "暫入贈/備品量", "暫入轉銷贈/備品量", "暫入歸還贈/備品量"
                        , "暫出贈/備品包裝量", "暫出轉銷贈/備品包裝量", "暫出歸還贈/備品包裝量", "暫入贈/備品包裝量", "暫入轉銷贈/備品包裝量", "暫入歸還贈/備品包裝量"
                        , "單位", "小單位", "結案碼", "包裝單位", "單價", "暫出金額", "暫入金額", "批號", "有效日期", "複檢日期", "來源單號", "備註"
                        , "專案代號", "專案名稱" };
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
                            string TF003 = item.TF003.ToString();
                            string DocNum = item.DocNum.ToString();
                            string TF024 = item.TF024.ToString();
                            string TF041 = item.TF041.ToString();
                            string TF005 = item.TF005.ToString();
                            string TF006 = item.TF006.ToString();
                            string TF007 = item.TF007.ToString();
                            string ME002 = item.ME002.ToString();
                            string TF008 = item.TF008.ToString();
                            string MV002 = item.MV002.ToString();
                            string TF009 = item.TF009.ToString();
                            string MB002 = item.MB002.ToString();
                            string TF011 = item.TF011.ToString();
                            string TaxType = item.TaxType.ToString();
                            string TF012 = item.TF012.ToString();
                            string TF013 = item.TF013.ToString();
                            string TF014 = item.TF014.ToString();
                            string TG004 = item.TG004.ToString();
                            string TG005 = item.TG005.ToString();
                            string TG006 = item.TG006.ToString();
                            string CatDept = item.CatDept.ToString();//责任部门
                            string MB006 = item.MB006.ToString() + " " + item.TypeTwoName02.ToString();//类别二
                            string MB007 = item.MB007.ToString() + " " + item.TypeTwoName03.ToString();//类别三
                            string TG007 = item.TG007.ToString() + " " + item.TempOut.ToString();
                            string TG008 = item.TG008.ToString() + " " + item.TypeTempIn.ToString();
                            string TG027 = item.TG027.ToString();
                            string TG035 = item.TG035.ToString();
                            string TG036 = item.TG036.ToString();
                            string TempOutQty = item.TempOutQty.ToString();
                            string TempOutSaleQty = item.TempOutSaleQty.ToString();
                            string TempOutReturnQty = item.TempOutReturnQty.ToString();
                            string TempInQty = item.TempInQty.ToString();
                            string TempInSaleQty = item.TempInSaleQty != null ? item.TempInSaleQty.ToString() : "";
                            string TempInReturnQty = item.TempInReturnQty != null ? item.TempInReturnQty.ToString() : "";
                            string TempOutPackQty = item.TempOutPackQty.ToString();
                            string TempOutPackSaleQty = item.TempOutPackSaleQty.ToString();
                            string TempOutPackReturnQty = item.TempOutPackReturnQty.ToString();
                            string TempInPackQty = item.TempInPackQty.ToString();
                            string TempInPackSaleQty = item.TempInPackSaleQty.ToString();
                            string TempInPackReturnQty = item.TempInPackReturnQty.ToString();
                            string TempOutSpareQty = item.TempOutSpareQty.ToString();
                            string TempOutSpareSaleQty = item.TempOutSpareSaleQty.ToString();
                            string TempOutSpareReturnQty = item.TempOutSpareReturnQty.ToString();
                            string TempInSpareQty = item.TempInSpareQty.ToString();
                            string TempInSpareSaleQty = item.TempInSpareSaleQty.ToString();
                            string TempInSpareReturnQty = item.TempInSpareReturnQty.ToString();
                            string TempOutSparePackQty = item.TempOutSparePackQty.ToString();
                            string TempOutSparePackSaleQty = item.TempOutSparePackSaleQty.ToString();
                            string TempOutSparePackReturnQty = item.TempOutSparePackReturnQty.ToString();
                            string TempInSparePackQty = item.TempInSparePackQty.ToString();
                            string TempInSparePackSaleQty = item.TempInSparePackSaleQty.ToString();
                            string TempInSparePackReturnQty = item.TempInSparePackReturnQty.ToString();
                            string TG010 = item.TG010.ToString();
                            string TG011 = item.TG011.ToString();
                            string CloseCode = item.CloseCode.ToString();
                            string TG031 = item.TG031.ToString();
                            string TG012 = item.TG012.ToString();
                            string TempOutAmount = item.TempOutAmount.ToString();
                            string TempInAmount = item.TempInAmount.ToString();
                            string TG017 = item.TG017.ToString();
                            string TG025 = item.TG025.ToString();
                            string TG026 = item.TG026.ToString();
                            string SourceDocNum = item.SourceDocNum.ToString();
                            string TG019 = item.TG019.ToString();
                            string TG018 = item.TG018.ToString();
                            string NB002 = item.NB002.ToString();
                            


                            var resultList = new[]
                            {
                                new {
                                    TF003,
                                    DocNum,
                                    TF024,
                                    TF041,
                                    TF005,
                                    TF006,
                                    TF007,
                                    ME002,
                                    TF008,
                                    MV002,
                                    TF009,
                                    MB002,
                                    TF011,
                                    TaxType,
                                    TF012,
                                    TF013,
                                    TF014,
                                    TG004,
                                    TG005,
                                    TG006,
                                    CatDept,
                                    MB006,
                                    MB007,
                                    TG007,
                                    TG008,
                                    TG027,
                                    TG035,
                                    TG036,
                                    TempOutQty,
                                    TempOutSaleQty,
                                    TempOutReturnQty,
                                    TempInQty,
                                    TempInSaleQty,
                                    TempInReturnQty,
                                    TempOutPackQty,
                                    TempOutPackSaleQty,
                                    TempOutPackReturnQty,
                                    TempInPackQty,
                                    TempInPackSaleQty,
                                    TempInPackReturnQty,
                                    TempOutSpareQty,
                                    TempOutSpareSaleQty,
                                    TempOutSpareReturnQty,
                                    TempInSpareQty,
                                    TempInSpareSaleQty,
                                    TempInSpareReturnQty,
                                    TempOutSparePackQty,
                                    TempOutSparePackSaleQty,
                                    TempOutSparePackReturnQty,
                                    TempInSparePackQty,
                                    TempInSparePackSaleQty,
                                    TempInSparePackReturnQty,
                                    TG010,
                                    TG011,
                                    CloseCode,
                                    TG031,
                                    TG012,
                                    TempOutAmount,
                                    TempInAmount,
                                    TG017,
                                    TG025,
                                    TG026,
                                    SourceDocNum,
                                    TG019,
                                    TG018,
                                    NB002,
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

        #region //ExcelTempOutReturnRecordDetail ERP暫出歸還單明細表輸出Excel
        public void ExcelTempOutReturnRecordDetail(string StardDay = "", string EndDay = "")
        {
            try
            {
                WebApiLoginCheck("MoldingConditions", "read,excel");
                List<string> OsrIdList = new List<string>();

                #region //Request
                dataRequest = pdmBasicInformationDA.GetTempOutReturnRecordDetail(StardDay, EndDay);
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
                    string excelFileName = "【MES2.0】暫出歸還單明細表";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "異動日期", "異動單號", "單據日期", "對象代碼", "對象簡稱", "部門代號", "簡稱", "員工代號", "簡稱"
                        , "廠別代號", "簡稱", "幣別", "課稅別", "匯率", "件數", "品號", "品名", "規格" , "責任部門", "類別二", "類別三", "轉出庫別", "轉入庫別"
                        , "轉出儲位", "轉入儲位", "暫出歸還數量", "暫出歸還包裝數量", "暫出歸還贈/備品量", "暫出歸還贈/備品包裝量"
                        , "單位", "小單位", "包裝單位", "單價", "暫出歸還金額", "批號", "有效日期", "複檢日期", "來源單號", "備註"
                        , "專案代號", "專案名稱" };
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
                            string TH003 = item.TH003.ToString();
                            string DocNum = item.DocNum.ToString();
                            string TH023 = item.TH023.ToString();
                            string TH005 = item.TH005.ToString();
                            string TH006 = item.TH006.ToString();
                            string TH007 = item.TH007.ToString();
                            string ME002 = item.ME002.ToString();
                            string TH008 = item.TH008.ToString();
                            string MV002 = item.MV002.ToString();
                            string TH009 = item.TH009.ToString();
                            string MB002 = item.MB002.ToString();
                            string TH011 = item.TH011.ToString();
                            string TH010 = item.TH010.ToString();
                            string TH012 = item.TH012.ToString();
                            string TH013 = item.TH013.ToString();
                            string TI004 = item.TI004.ToString();
                            string TI005 = item.TI005.ToString();
                            string TI006 = item.TI006.ToString();
                            string CatDept = item.CatDept.ToString();
                            string MB006 = item.MB006.ToString() + " " + item.TypeTwoName02.ToString();
                            string MB007 = item.TI008.ToString() + " " + item.TypeTwoName03.ToString();
                            string TI007 = item.TI007.ToString() + " " + item.TempOut.ToString();
                            string TI008 = item.TI008.ToString() + " " + item.TypeTempIn.ToString();
                            string TI026 = item.TI026.ToString();
                            string TI027 = item.TI027.ToString();
                            string TI009 = item.TI009.ToString();
                            string TI023 = item.TI023.ToString();
                            string TI035 = item.TI035.ToString();
                            string TI036 = item.TI036.ToString();
                            string TI010 = item.TI010.ToString();
                            string TI011 = item.TI011.ToString();
                            string TI024 = item.TI024.ToString();
                            string TI012 = item.TI012.ToString();
                            string TI013 = item.TI013.ToString();
                            string TI017 = item.TI017.ToString();
                            string TI018 = item.TI018.ToString();
                            string TI019 = item.TI019.ToString();
                            string SourceDocNum = item.SourceDocNum.ToString();
                            string TI021 = item.TI021.ToString();
                            string TI020 = item.TI020.ToString();

                            
                            var resultList = new[]
                            {
                                new {
                                TH003,
                                DocNum,
                                TH023,
                                TH005,
                                TH006,
                                TH007,
                                ME002,
                                TH008,
                                MV002,
                                TH009,
                                MB002,
                                TH011,
                                TH010,
                                TH012,
                                TH013,
                                TI004,
                                TI005,
                                TI006,
                                CatDept,
                                MB006,
                                MB007,
                                TI007,
                                TI008,
                                TI026,
                                TI027,
                                TI009,
                                TI023,
                                TI035,
                                TI036,
                                TI010,
                                TI011,
                                TI024,
                                TI012,
                                TI013,
                                TI017,
                                TI018,
                                TI019,
                                SourceDocNum,
                                TI021,
                                TI020,
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

    }
}