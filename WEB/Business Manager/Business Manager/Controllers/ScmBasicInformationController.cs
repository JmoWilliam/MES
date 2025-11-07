using Helpers;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ScmBasicInformationController : WebController
    {
        private ScmBasicInformationDA scmBasicInformationDA = new ScmBasicInformationDA();

        #region //View
        public ActionResult PackingVolume()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PackingWeight()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ItemDefaultWeight()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult CustomerManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ForwarderManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DeliveryCustomer()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SupplierManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult InventoryManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult Packing()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RfqProductClassManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductUseManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult EcRegisterManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PerformanceGoalsManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductTypeGroupManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductTypeManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QuotationTagManagement()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult SupplierProcessEquipmentReport()
        {
            //參與托外掃碼的供應商已登記的製程與機台
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProjectBudgetManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProjectDetail()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetPackingVolume 取得包材體積資料
        [HttpPost]
        public void GetPackingVolume(int VolumeId = -1, string VolumeNo = "", string VolumeName = "", string VolumeType = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PackingVolume", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetPackingVolume(VolumeId, VolumeNo, VolumeName, VolumeType, Status
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

        #region //GetPackingWeight 取得包材重量資料
        [HttpPost]
        public void GetPackingWeight(int WeightId = -1, string WeightNo = "", string WeightName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PackingWeight", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetPackingWeight(WeightId, WeightNo, WeightName, Status
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

        #region //GetItemWeight 取得物件重量資料
        [HttpPost]
        public void GetItemWeight(int ItemDefaultWeightId = -1, int MtlModelId = -1, string Status = "", int StartParent = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ItemDefaultWeight", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetItemWeight(ItemDefaultWeightId, MtlModelId, Status, StartParent
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

        #region //GetCustomer 取得客戶資料
        [HttpPost]
        public void GetCustomer(int CustomerId = -1, string CustomerNo = "", string CustomerName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetCustomer(CustomerId, CustomerNo, CustomerName, Status
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

        #region //GetForwarder 取得貨運承攬商資料
        [HttpPost]
        public void GetForwarder(int ForwarderId = -1, string ShipMethod = "", string ForwarderNo = "", string ForwarderName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ForwarderManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetForwarder(ForwarderId, ShipMethod, ForwarderNo, ForwarderName, Status
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

        #region //GetDeliveryCustomer 取得送貨客戶資料
        [HttpPost]
        public void GetDeliveryCustomer(int DcId = -1, int CustomerId = -1, string DcName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeliveryCustomer", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetDeliveryCustomer(DcId, CustomerId, DcName, Status
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

        #region //GetSupplier 取得供應商資料
        [HttpPost]
        public void GetSupplier(int SupplierId = -1, string SupplierNo = "", string SupplierName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetSupplier(SupplierId, SupplierNo, SupplierName, Status
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

        #region //GetInventory 取得庫別資料
        [HttpPost]
        public void GetInventory(int InventoryId = -1, string InventoryNo = "", string InventoryName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetInventory(InventoryId, InventoryNo, InventoryName, Status
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

        #region //GetPacking 取得包材資料
        [HttpPost]
        public void GetPacking(int PackingId = -1, string PackingName = "", string PackingType = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("Packing", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetPacking(PackingId, PackingName, PackingType, Status
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

        #region //GetRfqProductClass -- 取得RFQ產品類型 -- Chia Yuan 2023.06.30
        [HttpPost]
        public void GetRfqProductClass(int RfqProClassId = -1, string RfqProductClassName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetRfqProductClass(RfqProClassId, RfqProductClassName, Status
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

        #region //GetRfqProductType -- 取得RFQ產品類別 -- Chia Yuan 2023.07.03
        [HttpPost]
        public void GetRfqProductType(int RfqProTypeId = -1, int RfqProClassId = -1, string RfqProductTypeName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetRfqProductType(RfqProTypeId, RfqProClassId, RfqProductTypeName, Status
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

        #region //GetRfqPackageType -- 取得RFQ包裝種類 -- Chia Yuan 2023.07.03
        [HttpPost]
        public void GetRfqPackageType(int RfqPkTypeId = -1, int RfqProClassId = -1, string PackagingMethod = "", string Status = "", string SustSupplyStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetRfqPackageType(RfqPkTypeId, RfqProClassId, PackagingMethod, Status, SustSupplyStatus
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

        #region GetProductUse -- 取得RFQ產品用途 -- Chia Yuan 2023.07.11
        [HttpPost]
        public void GetProductUse(int ProductUseId = -1, string ProductUseNo = "", string ProductUseName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductUseManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProductUse(ProductUseId, ProductUseNo, ProductUseName, Status
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

        #region GetMember -- 取得電商註冊資訊
        [HttpPost]
        public void GetMember(int MemberId = -1, string MemberName = "", string OrgShortName = "", string OrganizaitonType = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("EcRegisterManagment", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetMember(MemberId, MemberName, OrgShortName, OrganizaitonType, Status
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

        #region //GetPerformanceGoals 取得業務績效目標
        [HttpPost]
        public void GetPerformanceGoals(int PgId = -1, string PgNo = "", string PgName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("PerformanceGoalsManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetPerformanceGoals(PgId, PgNo, PgName, Status
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

        #region //GetPgDetail 取得業務績效目標單身
        [HttpPost]
        public void GetPgDetail(int PgId = -1, int PgDetailId = -1, int UserId = -1, string IntentType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetPgDetail(PgId, PgDetailId, UserId, IntentType
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

        #region //GetProductTypeGroup 取得產品群組資料
        [HttpPost]
        public void GetProductTypeGroup(int ProTypeGroupId = -1, int RfqProClassId = -1, string ProTypeGroupName = "", string CoatingFlag ="", string Status =""
            , string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductTypeGroupManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProductTypeGroup(ProTypeGroupId, RfqProClassId, ProTypeGroupName, CoatingFlag, Status, SearchKey
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

        #region //GetTemplateRfiSignFlow 取得市場評估單流程
        [HttpPost]
        public void GetTemplateRfiSignFlow(int TempRfiSfId = -1, int ProTypeGroupId = -1, int DepartmentId = -1, int FlowUser = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetTemplateRfiSignFlow(TempRfiSfId, ProTypeGroupId, DepartmentId, FlowUser, Status
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

        #region //GetTemplateDesignSignFlow 取得設計申請單流程
        [HttpPost]
        public void GetTemplateDesignSignFlow(int TempDesignSfId = -1, int ProTypeGroupId = -1, int DepartmentId = -1, int FlowUser = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetTemplateDesignSignFlow(TempDesignSfId, ProTypeGroupId, DepartmentId, FlowUser, Status
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

        #region //GetProductType 取得產品群組資料
        [HttpPost]
        public void GetProductType(int ProTypeId = -1, int ProTypeGroupId = -1,string ProTypeName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductTypeManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProductType(ProTypeId, ProTypeGroupId, ProTypeName,  Status
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

        #region //GetQuotationTag 取得報價單屬性標籤
        [HttpPost]
        public void GetQuotationTag(int QtId = -1, string TagNo = "", string TagName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QuotationTagManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetQuotationTag(QtId, TagNo, TagName, Status
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

        #region//GetSupplierProcessEquipment --參與托外掃碼的供應商已登記的製程與機台清單
        [HttpPost]
        public void GetSupplierProcessEquipment(int SmId = -1, int SupplierId = -1, int ProcessId = -1, int MachineId = -1, int ShopId = -1
        , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetSupplierProcessEquipment(SmId, SupplierId, ProcessId, MachineId, ShopId
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
        public void GetProcess(int SupplierId = -1, int ProcessId = -1, string ProcessNo = "", string ProcessName = ""
            , string Status = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProcess(SupplierId, ProcessId, ProcessNo, ProcessName, Status
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

        #region //GetWorkShop 取得車間資料
        [HttpPost]
        public void GetWorkShop(int SupplierId = -1, int ShopId = -1, string ShopNo = "", string ShopName = ""
            , string Status = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetWorkShop(SupplierId, ShopId, ShopNo, ShopName, Status
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

        #region //GetMachine 取得機台資料
        [HttpPost]
        public void GetMachine(int SupplierId = -1, int ShopId = -1, int MachineId = -1, string MachineNo = "", string MachineName = ""
            , string Status = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetMachine(SupplierId, ShopId, MachineId, MachineNo, MachineName, Status
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

        #region //GetProcessItem 取得托外製程相關資料
        [HttpPost]
        public void GetProcessItem(int ProcessId = -1, string ProcessNo = "", string ProcessName = "", string ProcessDesc = ""
            , string Status = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProcessItem(ProcessId, ProcessNo, ProcessName, ProcessDesc, Status
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
        
        #region //GetMachineItem 取得托外機台相關資料
        [HttpPost]
        public void GetMachineItem(int ShopId = -1, string ShopName = "", int MachineId = -1, string MachineName = ""
            , string Status = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "read,constrained-data");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetMachineItem(ShopId, ShopName, MachineId, MachineName, Status
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

        #region //GetProject 取得專案資料
        [HttpPost]
        public void GetProject(int ProjectId = -1, string ProjectNo = "", string ProjectName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProject(ProjectId, ProjectNo, ProjectName
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

        #region //GetProjectDetail 取得專案詳細資料
        [HttpPost]
        public void GetProjectDetail(int ProjectDetailId = -1, int ProjectId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProjectDetail(ProjectDetailId, ProjectId
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

        #region //GetProjectBudgetChangeLog 取得專案預算變更Log紀錄
        [HttpPost]
        public void GetProjectBudgetChangeLog(int ProjectDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "read");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProjectBudgetChangeLog(ProjectDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddPackingVolume 包材體積資料新增
        [HttpPost]
        public void AddPackingVolume(int VolumeId = -1, string VolumeNo = "", string VolumeName = "", string VolumeType = "", string VolumeSpec = "", double Volume = 0.0, int VolumeUomId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingVolume", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddPackingVolume(VolumeId, VolumeNo, VolumeName, VolumeType, VolumeSpec, Volume, VolumeUomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPackingWeight 包材重量資料新增
        [HttpPost]
        public void AddPackingWeight(int WeightId = -1, string WeightNo = "", string WeightName = "", string WeightSpec = "", double Weight = 0.0, int WeightUomId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingWeight", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddPackingWeight(WeightId, WeightNo, WeightName, WeightSpec, Weight, WeightUomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddItemWeight 物件重量資料新增
        [HttpPost]
        public void AddItemWeight(int ItemDefaultWeightId = -1, int MtlModelId = -1, string DefaultWeight = "", int DefaultWeightUomId = -1, string Status = "", int ParentId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemDefaultWeight", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddItemWeight(ItemDefaultWeightId, MtlModelId, DefaultWeight, DefaultWeightUomId, Status, ParentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddCustomer 客戶資料新增
        [HttpPost]
        public void AddCustomer(int CustomerId = -1, string CustomerNo = "", string CustomerName = "", string CustomerEnglishName = "", string CustomerShortName = ""
            , string RelatedPerson = "", string PermitDate = "", string Version = "", string ResponsiblePerson = "", string Contact = "", string TelNoFirst = ""
            , string TelNoSecond = "", string FaxNo = "", string Email = "", string GuiNumber = "", double Capital = 0.0, double AnnualTurnover = 0.0, int Headcount = -1
            , string HomeOffice = "", string Currency = "", int DepartmentId = -1, string CustomerKind = "", int SalesmenId = -1, int PaymentSalesmenId = -1
            , string InauguateDate = "", string CloseDate = "", string ZipCodeRegister = "", string RegisterAddressFirst = "", string RegisterAddressSecond = ""
            , string ZipCodeInvoice = "", string InvoiceAddressFirst = "", string InvoiceAddressSecond = "", string ZipCodeDelivery = "", string DeliveryAddressFirst = ""
            , string DeliveryAddressSecond = "", string ZipCodeDocument = "", string DocumentAddressFirst = "", string DocumentAddressSecond = "", string BillReceipient = ""
            , string ZipCodeBill = "", string BillAddressFirst = "", string BillAddressSecond = "", string InvocieAttachedStatus = "", double DepositRate = 0.0
            , string TaxAmountCalculateType = "", string SaleRating = "", string CreditRating = "", string TradeTerm = "", string PaymentTerm = "", string PricingType = ""
            , string ClearanceType = "", string DocumentDeliver = "", string ReceiptReceive = "", string PaymentType = "", string TaxNo = "", string InvoiceCount = ""
            , string Taxation = "", string Country = "", string Region = "", string Route = "", string UploadType = "", string PaymentBankFirst = "", string BankAccountFirst = ""
            , string PaymentBankSecond = "", string BankAccountSecond = "", string PaymentBankThird = "", string BankAccountThird = "", string Account = ""
            , string AccountInvoice = "", string AccountDay = "", string ShipMethod = "", string ShipType = "", int ForwarderId = -1, string CustomerRemark = ""
            , double CreditLimit = 0.0, string CreditLimitControl = "", string CreditLimitControlCurrency = "", string SoCreditAuditType = "", string SiCreditAuditType = ""
            , string DoCreditAuditType = "", string InTransitCreditAuditType = "", string TransferStatus = "", string TransferDate = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("CustomerManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddCustomer(CustomerId, CustomerNo, CustomerName, CustomerEnglishName
                    , CustomerShortName, RelatedPerson, PermitDate, Version, ResponsiblePerson, Contact, TelNoFirst, TelNoSecond
                    , FaxNo, Email, GuiNumber, Capital, AnnualTurnover, Headcount, HomeOffice, Currency, DepartmentId, CustomerKind
                    , SalesmenId, PaymentSalesmenId, InauguateDate, CloseDate, ZipCodeRegister, RegisterAddressFirst, RegisterAddressSecond
                    , ZipCodeInvoice, InvoiceAddressFirst, InvoiceAddressSecond, ZipCodeDelivery, DeliveryAddressFirst, DeliveryAddressSecond
                    , ZipCodeDocument, DocumentAddressFirst, DocumentAddressSecond, BillReceipient, ZipCodeBill, BillAddressFirst, BillAddressSecond
                    , InvocieAttachedStatus, DepositRate, TaxAmountCalculateType, SaleRating, CreditRating, TradeTerm, PaymentTerm, PricingType
                    , ClearanceType, DocumentDeliver, ReceiptReceive, PaymentType, TaxNo, InvoiceCount, Taxation, Country
                    , Region, Route, UploadType, PaymentBankFirst, BankAccountFirst, PaymentBankSecond, BankAccountSecond
                    , PaymentBankThird, BankAccountThird, Account, AccountInvoice, AccountDay, ShipMethod, ShipType, ForwarderId
                    , CustomerRemark, CreditLimit, CreditLimitControl, CreditLimitControlCurrency, SoCreditAuditType, SiCreditAuditType
                    , DoCreditAuditType, InTransitCreditAuditType, TransferStatus, TransferDate, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddForwarder 貨運承攬商資料新增
        [HttpPost]
        public void AddForwarder(int ForwarderId = -1, string ShipMethod = "", string ForwarderNo = "", string ForwarderName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("CustomerManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddForwarder(ForwarderId, ShipMethod, ForwarderNo, ForwarderName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeliveryCustomer 送貨客戶資料新增
        [HttpPost]
        public void AddDeliveryCustomer(int DcId = -1, int CustomerId = -1, string DcName = "", string DcEnglishName = "", string DcShortName = "", string Contact = ""
            , string TelNo = "", string FaxNo = "", string RegisteredAddress = "", string DeliveryAddress = "", string ShipType = "", int ForwarderId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryCustomer", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddDeliveryCustomer(DcId, CustomerId, DcName, DcEnglishName, DcShortName, Contact
                    , TelNo, FaxNo, RegisteredAddress, DeliveryAddress, ShipType, ForwarderId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddSupplier 供應商資料新增
        [HttpPost]
        public void AddSupplier(int SupplierId = -1, string SupplierNo = "", string SupplierName = "", string SupplierShortName = "", string SupplierEnglishName = "", string Country = ""
            , string Region = "", string GuiNumber = "", string ResponsiblePerson = "", string ContactFirst = "", string ContactSecond = "", string ContactThird = "", string TelNoFirst = "", string TelNoSecond = ""
            , string FaxNo = "", string FaxNoAccounting = "", string Email = "", string AddressFirst = "", string AddressSecond = "", string ZipCodeFirst = "", string ZipCodeSecond = ""
            , string BillAddressFirst = "", string BillAddressSecond = "", string PermitStatus = "", string InauguateDate = "", string AccountMonth = "", string AccountDay = ""
            , string Version = "", double Capital = 0.0, int Headcount = -1, string PoDeliver = "", string Currency = "", string TradeTerm = "", string PaymentType = "", string PaymentTerm = ""
            , string ReceiptReceive = "", string InvoiceCount = "", string TaxNo = "", string Taxation = "", string PermitPartialDelivery = "", string TaxAmountCalculateType = "", string InvocieAttachedStatus = ""
            , string CertificateFormatType = "", double DepositRate = 0.0, string TradeItem = "", string RemitBank = "", string RemitAccount = "", string AccountPayable = "", string AccountOverhead = ""
            , string AccountInvoice = "", string SupplierLevel = "", string DeliveryRating = "", string QualityRating = "", string RelatedPerson = "", int PurchaseUserId = -1, string SupplierRemark = ""
            , string PassStationControl ="", string TransferStatus = "", string TransferDate  = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddSupplier(SupplierId, SupplierNo, SupplierName, SupplierShortName, SupplierEnglishName, Country
                    , Region, GuiNumber, ResponsiblePerson, ContactFirst, ContactSecond, ContactThird, TelNoFirst, TelNoSecond
                    , FaxNo, FaxNoAccounting, Email, AddressFirst, AddressSecond, ZipCodeFirst , ZipCodeSecond
                    , BillAddressFirst, BillAddressSecond, PermitStatus, InauguateDate, AccountMonth, AccountDay
                    , Version, Capital, Headcount, PoDeliver, Currency, TradeTerm, PaymentType, PaymentTerm
                    , ReceiptReceive, InvoiceCount, TaxNo, Taxation, PermitPartialDelivery,TaxAmountCalculateType, InvocieAttachedStatus
                    , CertificateFormatType, DepositRate, TradeItem, RemitBank, RemitAccount, AccountPayable, AccountOverhead
                    , AccountInvoice, SupplierLevel, DeliveryRating, QualityRating, RelatedPerson, PurchaseUserId, SupplierRemark
                    , PassStationControl, TransferStatus, TransferDate, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddInventory 庫別資料新增
        [HttpPost]
        public void AddInventory(int InventoryId = -1, string InventoryNo = "", string InventoryName = "", string InventoryType = ""
            , string MrpCalculation = "", string ConfirmStatus = "", string SaveStatus = "", string InventoryDesc = ""
            , string TransferStatus = "", string TransferDate = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("InventoryManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddInventory(InventoryId, InventoryNo, InventoryName, InventoryType
                    , MrpCalculation, ConfirmStatus, SaveStatus, InventoryDesc, TransferStatus, TransferDate, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPacking 包材資料新增
        [HttpPost]
        public void AddPacking(int PackingId = -1, string PackingName = "", string PackingType = "", string VolumeSpec = "", double Volume = 0.0
             , int VolumeUomId = -1, string WeightSpec = "", double Weight = 0.0, int WeightUomId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("Packing", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddPacking(PackingId, PackingName, PackingType, VolumeSpec
                    , Volume, VolumeUomId, WeightSpec, Weight, WeightUomId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRfqProductClass -- RFQ產品類型新增 -- Chia Yuan 2023.07.03
        [HttpPost, ValidateAntiForgeryToken]
        public void AddRfqProductClass(int RfqProClassId = -1, string RfqProductClassName = null)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddRfqProductClass(RfqProClassId, RfqProductClassName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRfqProductType -- RFQ產品類別新增 -- Chia Yuan 2023.07.03
        [HttpPost, ValidateAntiForgeryToken]
        public void AddRfqProductType(int RfqProTypeId = -1, int RfqProClassId = -1, string RfqProductTypeName = null)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "rfqproduct-type");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddRfqProductType(RfqProTypeId, RfqProClassId, RfqProductTypeName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRfqPackageType -- RFQ包裝種類新增 -- Chia Yuan 2023.07.03
        [HttpPost, ValidateAntiForgeryToken]
        public void AddRfqPackageType(int RfqPkTypeId = -1, int RfqProClassId = -1, string PackagingMethod = "", string SustSupplyStatus = "")
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "rfqproduct-package");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddRfqPackageType(RfqPkTypeId, RfqProClassId, PackagingMethod, SustSupplyStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProductUse RFQ產品用途新增
        [HttpPost]
        public void AddProductUse(int ProductUseId = -1, string ProductUseNo = "", string ProductUseName = "", string TypeOne = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProductUseManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddProductUse(ProductUseId, ProductUseNo, ProductUseName, TypeOne, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPerformanceGoals 業務績效目標新增
        [HttpPost]
        public void AddPerformanceGoals(string PgNo = "", string PgName = "", string PgDesc = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddPerformanceGoals(PgNo, PgName, PgDesc, StartDate, EndDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPgDetail 新增業務績效目標單身
        [HttpPost]
        public void AddPgDetail(int PgId = -1,string IntentType = "", int UserId = -1, int IntentAmount = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddPgDetail(PgId, IntentType, UserId, IntentAmount);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProductTypeGroup 新增產品群組資料
        [HttpPost]
        public void AddProductTypeGroup(int RfqProClassId = -1, string TypeOne = "", string ProTypeGroupName = "", string CoatingFlag = "")
        {
            try
            {
                WebApiLoginCheck("ProductTypeGroupManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddProductTypeGroup(RfqProClassId, TypeOne, ProTypeGroupName, CoatingFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //AddProductType 新增產品類別資料
        [HttpPost]
        public void AddProductType(int ProTypeGroupId = -1, string ProTypeName = "")
        {
            try
            {
                WebApiLoginCheck("ProductTypeManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddProductType(ProTypeGroupId, ProTypeName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //AddQuotationTag 新增報價單屬性標籤
        [HttpPost]
        public void AddQuotationTag(string TagNo = "", string TagName = "")
        {
            try
            {
                WebApiLoginCheck("QuotationTagManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddQuotationTag(TagNo, TagName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddSupplierProcessEquipment -- 參與托外掃碼的供應商的製程與機台清單新增
        [HttpPost]
        public void AddSupplierProcessEquipment(int SupplierId = -1, int ProcessId = -1, int MachineId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddSupplierProcessEquipment(SupplierId, ProcessId, MachineId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProject 新增專案資料
        [HttpPost]
        public void AddProject(string ProjectNo = "", string ProjectName = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddProject(ProjectNo, ProjectName, EffectiveDate, ExpirationDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectDetail 新增專案預算資料
        [HttpPost]
        public void AddProjectDetail(int ProjectId = -1, string ProjectType = "", string Currency = "", double ExchangeRate = -1, double BudgetAmount = -1, double LocalBudgetAmount = -1, string Remark = "", string ProjectFile = "")
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddProjectDetail(ProjectId, ProjectType, Currency, ExchangeRate, BudgetAmount, LocalBudgetAmount, Remark, ProjectFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectBudgetChangeLog 新增專案預算變更資料
        [HttpPost]
        public void AddProjectBudgetChangeLog(int ProjectDetailId = -1, string Currency = "", double ExchangeRate = -1, double BudgetAmount = -1, double LocalBudgetAmount = -1, string Remark = "", string ProjectFile = "")
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddProjectBudgetChangeLog(ProjectDetailId, Currency, ExchangeRate, BudgetAmount, LocalBudgetAmount, Remark, ProjectFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdatePackingVolume 包材體積資料更新
        //[HttpPost]
        //public void UpdatePackingVolume(int VolumeId = -1, string VolumeNo = "", string VolumeName = "", string VolumeType = "", string VolumeSpec = "", double Volume = 0.0, int VolumeUomId = -1)
        //{
        //    try
        //    {
        //        WebApiLoginCheck("PackingVolume", "update");

        //        #region //Request
        //        scmBasicInformationDA = new ScmBasicInformationDA();
        //        dataRequest = scmBasicInformationDA.UpdatePackingVolume(VolumeId, VolumeNo, VolumeName, VolumeType, VolumeSpec, Volume, VolumeUomId);
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

        #region //UpdateVolumeStatus 包材體積狀態更新
        [HttpPost]
        public void UpdateVolumeStatus(int VolumeId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingVolume", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateVolumeStatus(VolumeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePackingWeight 包材重量資料更新
        [HttpPost]
        public void UpdatePackingWeight(int WeightId = -1, string WeightNo = "", string WeightName = "", string WeightSpec = "", double Weight = 0.0, int WeightUomId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingWeight", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePackingWeight(WeightId, WeightNo, WeightName, WeightSpec, Weight, WeightUomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWeightStatus 包材重量狀態更新
        [HttpPost]
        public void UpdateWeightStatus(int WeightId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingWeight", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateWeightStatus(WeightId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemWeight 物件重量資料更新
        [HttpPost]
        public void UpdateItemWeight(int ItemDefaultWeightId = -1, int MtlModelId = -1, string DefaultWeight = "", int DefaultWeightUomId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemDefaultWeight", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateItemWeight(ItemDefaultWeightId, MtlModelId, DefaultWeight, DefaultWeightUomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItemWeightStatus 物件重量狀態更新
        [HttpPost]
        public void UpdateItemWeightStatus(int ItemDefaultWeightId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemDefaultWeight", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateItemWeightStatus(ItemDefaultWeightId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCustomer 客戶資料更新
        [HttpPost]
        public void UpdateCustomer(int CustomerId = -1, string CustomerNo = "", string CustomerName = "", string CustomerEnglishName = "", string CustomerShortName = ""
            , string RelatedPerson = "", string PermitDate = "", string Version = "", string ResponsiblePerson = "", string Contact = "", string TelNoFirst = ""
            , string TelNoSecond = "", string FaxNo = "", string Email = "", string GuiNumber = "", double Capital = 0.0, double AnnualTurnover = 0.0, int Headcount = -1
            , string HomeOffice = "", string Currency = "", int DepartmentId = -1, string CustomerKind = "", int SalesmenId = -1, int PaymentSalesmenId = -1
            , string InauguateDate = "", string CloseDate = "", string ZipCodeRegister = "", string RegisterAddressFirst = "", string RegisterAddressSecond = ""
            , string ZipCodeInvoice = "", string InvoiceAddressFirst = "", string InvoiceAddressSecond = "", string ZipCodeDelivery = "", string DeliveryAddressFirst = ""
            , string DeliveryAddressSecond = "", string ZipCodeDocument = "", string DocumentAddressFirst = "", string DocumentAddressSecond = "", string BillReceipient = ""
            , string ZipCodeBill = "", string BillAddressFirst = "", string BillAddressSecond = "", string InvocieAttachedStatus = "", double DepositRate = 0.0
            , string TaxAmountCalculateType = "", string SaleRating = "", string CreditRating = "", string TradeTerm = "", string PaymentTerm = "", string PricingType = ""
            , string ClearanceType = "", string DocumentDeliver = "", string ReceiptReceive = "", string PaymentType = "", string TaxNo = "", string InvoiceCount = ""
            , string Taxation = "", string Country = "", string Region = "", string Route = "", string UploadType = "", string PaymentBankFirst = "", string BankAccountFirst = ""
            , string PaymentBankSecond = "", string BankAccountSecond = "", string PaymentBankThird = "", string BankAccountThird = "", string Account = "", string AccountInvoice = ""
            , string AccountDay = "", string ShipMethod = "", string ShipType = "", int ForwarderId = -1, string CustomerRemark = "", double CreditLimit = 0.0
            , string CreditLimitControl = "", string CreditLimitControlCurrency = "", string SoCreditAuditType = "", string SiCreditAuditType = ""
            , string DoCreditAuditType = "", string InTransitCreditAuditType = "", string TransferStatus = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("CustomerManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateCustomer(CustomerId, CustomerNo, CustomerName, CustomerEnglishName, CustomerShortName
                    , RelatedPerson, PermitDate, Version, ResponsiblePerson, Contact, TelNoFirst, TelNoSecond, FaxNo, Email
                    , GuiNumber, Capital, AnnualTurnover, Headcount, HomeOffice, Currency, DepartmentId, CustomerKind
                    , SalesmenId, PaymentSalesmenId, InauguateDate, CloseDate, ZipCodeRegister, RegisterAddressFirst, RegisterAddressSecond
                    , ZipCodeInvoice, InvoiceAddressFirst, InvoiceAddressSecond, ZipCodeDelivery, DeliveryAddressFirst, DeliveryAddressSecond
                    , ZipCodeDocument, DocumentAddressFirst, DocumentAddressSecond, BillReceipient, ZipCodeBill, BillAddressFirst, BillAddressSecond
                    , InvocieAttachedStatus, DepositRate, TaxAmountCalculateType, SaleRating, CreditRating, TradeTerm, PaymentTerm, PricingType
                    , ClearanceType, DocumentDeliver, ReceiptReceive, PaymentType, TaxNo, InvoiceCount, Taxation, Country
                    , Region, Route, UploadType, PaymentBankFirst, BankAccountFirst, PaymentBankSecond, BankAccountSecond
                    , PaymentBankThird, BankAccountThird, Account, AccountInvoice, AccountDay, ShipMethod, ShipType, ForwarderId
                    , CustomerRemark, CreditLimit, CreditLimitControl, CreditLimitControlCurrency, SoCreditAuditType, SiCreditAuditType
                    , DoCreditAuditType, InTransitCreditAuditType, TransferStatus, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCustomerStatus 客戶狀態更新
        [HttpPost]
        public void UpdateCustomerStatus(int CustomerId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateCustomerStatus(CustomerId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateForwarder 貨運承攬商資料更新
        [HttpPost]
        public void UpdateForwarder(int ForwarderId = -1, string ShipMethod = "", string ForwarderNo = "", string ForwarderName = "")
        {
            try
            {
                WebApiLoginCheck("ForwarderManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateForwarder(ForwarderId, ShipMethod, ForwarderNo, ForwarderName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateForwarderStatus 貨運承攬商狀態更新
        [HttpPost]
        public void UpdateForwarderStatus(int ForwarderId = -1)
        {
            try
            {
                WebApiLoginCheck("ForwarderManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateForwarderStatus(ForwarderId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDeliveryCustomer 送貨客戶資料更新
        [HttpPost]
        public void UpdateDeliveryCustomer(int DcId = -1, int CustomerId = -1, string DcName = "", string DcEnglishName = "", string DcShortName = "", string Contact = ""
            , string TelNo = "", string FaxNo = "", string RegisteredAddress = "", string DeliveryAddress = "", string ShipType = "", int ForwarderId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("DeliveryCustomer", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateDeliveryCustomer(DcId, CustomerId, DcName, DcEnglishName, DcShortName, Contact
                    , TelNo, FaxNo, RegisteredAddress, DeliveryAddress, ShipType, ForwarderId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDeliveryCustomerStatus 送貨客戶狀態更新
        [HttpPost]
        public void UpdateDeliveryCustomerStatus(int DcId = -1)
        {
            try
            {
                WebApiLoginCheck("DeliveryCustomer", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateDeliveryCustomerStatus(DcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSupplier 供應商資料更新
        [HttpPost]
        public void UpdateSupplier(int SupplierId = -1, string SupplierNo = "", string SupplierName = "", string SupplierShortName = "", string SupplierEnglishName = "", string Country = ""
            , string Region = "", string GuiNumber = "", string ResponsiblePerson = "", string ContactFirst = "", string ContactSecond = "", string ContactThird = "", string TelNoFirst = "", string TelNoSecond = ""
            , string FaxNo = "", string FaxNoAccounting = "", string Email = "", string AddressFirst = "", string AddressSecond = "", string ZipCodeFirst = "", string ZipCodeSecond = ""
            , string BillAddressFirst = "", string BillAddressSecond = "", string PermitStatus = "", string InauguateDate = "", string AccountMonth = "", string AccountDay = ""
            , string Version = "", double Capital = 0.0, int Headcount = -1, string PoDeliver = "", string Currency = "", string TradeTerm = "", string PaymentType = "", string PaymentTerm = ""
            , string ReceiptReceive = "", string InvoiceCount = "", string TaxNo = "", string Taxation = "", string PermitPartialDelivery = "", string TaxAmountCalculateType = "", string InvocieAttachedStatus = ""
            , string CertificateFormatType = "", double DepositRate = 0.0, string TradeItem = "", string RemitBank = "", string RemitAccount = "", string AccountPayable = "", string AccountOverhead = ""
            , string AccountInvoice = "", string SupplierLevel = "", string DeliveryRating = "", string QualityRating = "", string RelatedPerson = "", int PurchaseUserId = -1, string SupplierRemark = ""
            , string PassStationControl = "", string TransferStatus = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateSupplier(SupplierId, SupplierNo, SupplierName, SupplierShortName, SupplierEnglishName, Country
                    , Region, GuiNumber, ResponsiblePerson, ContactFirst, ContactSecond, ContactThird, TelNoFirst, TelNoSecond
                    , FaxNo, FaxNoAccounting, Email, AddressFirst, AddressSecond, ZipCodeFirst, ZipCodeSecond
                    , BillAddressFirst, BillAddressSecond, PermitStatus, InauguateDate, AccountMonth, AccountDay
                    , Version, Capital, Headcount, PoDeliver, Currency, TradeTerm, PaymentType, PaymentTerm
                    , ReceiptReceive, InvoiceCount, TaxNo, Taxation, PermitPartialDelivery, TaxAmountCalculateType, InvocieAttachedStatus
                    , CertificateFormatType, DepositRate, TradeItem, RemitBank, RemitAccount, AccountPayable, AccountOverhead
                    , AccountInvoice, SupplierLevel, DeliveryRating, QualityRating, RelatedPerson, PurchaseUserId, SupplierRemark
                    , PassStationControl, TransferStatus, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSupplierStatus 供應商狀態更新
        [HttpPost]
        public void UpdateSupplierStatus(int SupplierId = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateSupplierStatus(SupplierId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSupplierPassStationControl 供應商是否刷過站狀態更新
        [HttpPost]
        public void UpdateSupplierPassStationControl(int SupplierId = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateSupplierPassStationControl(SupplierId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateInventory 庫別資料更新
        [HttpPost]
        public void UpdateInventory(int InventoryId = -1, string InventoryNo = "", string InventoryName = "", string InventoryType = ""
            , string MrpCalculation = "", string ConfirmStatus = "", string SaveStatus = "", string InventoryDesc = ""
            , string TransferStatus = "", string TransferDate = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("InventoryManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateInventory(InventoryId, InventoryNo, InventoryName, InventoryType
                    , MrpCalculation, ConfirmStatus, SaveStatus, InventoryDesc, TransferStatus, TransferDate, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateInventoryStatus 庫別狀態更新
        [HttpPost]
        public void UpdateInventoryStatus(int InventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateInventoryStatus(InventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePacking 包材資料更新
        [HttpPost]
        public void UpdatePacking(int PackingId = -1, string PackingName = "", string PackingType = "", string VolumeSpec = "", double Volume = 0.0
             , int VolumeUomId = -1, string WeightSpec = "", double Weight = 0.0, int WeightUomId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("Packing", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePacking(PackingId, PackingName, PackingType, VolumeSpec
                    , Volume, VolumeUomId, WeightSpec, Weight, WeightUomId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePackingStatus 包材狀態更新
        [HttpPost]
        public void UpdatePackingStatus(int PackingId = -1)
        {
            try
            {
                WebApiLoginCheck("Packing", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePackingStatus(PackingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqProductClass -- RFQ產品類型更新 -- Chia Yuan 230630

        [HttpPost, ValidateAntiForgeryToken]
        public void UpdateRfqProductClass(int RfqProClassId = -1, string RfqProductClassName = null)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateRfqProductClass(RfqProClassId, RfqProductClassName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqProductClassStatus -- RFQ產品類型狀態更新 -- Chia Yuan 230630

        [HttpPost, ValidateAntiForgeryToken]
        public void UpdateRfqProductClassStatus(int RfqProClassId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateRfqProductClassStatus(RfqProClassId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateRfqProductType -- RFQ產品類別更新 -- Chia Yuan 230703

        [HttpPost, ValidateAntiForgeryToken]
        public void UpdateRfqProductType(int RfqProTypeId = -1, string RfqProductTypeName = "")
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "rfqproduct-type");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateRfqProductType(RfqProTypeId, RfqProductTypeName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqProductClassStatus -- RFQ產品類別狀態更新 -- Chia Yuan 230703

        [HttpPost, ValidateAntiForgeryToken]
        public void UpdateRfqProductTypeStatus(int RfqProTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateRfqProductTypeStatus(RfqProTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateRfqPackageType -- RFQ包裝種類更新 -- Chia Yuan 230703
        [HttpPost, ValidateAntiForgeryToken]
        public void UpdateRfqPackageType(int RfqPkTypeId = -1, string PackagingMethod = "", string SustSupplyStatus = "")
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "rfqproduct-package");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateRfqPackageType(RfqPkTypeId, PackagingMethod, SustSupplyStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqPackageTypeStatus -- RFQ包裝種類狀態更新 -- Chia Yuan 230703
        [HttpPost, ValidateAntiForgeryToken]
        public void UpdateRfqPackageTypeStatus(int RfqPkTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateRfqPackageTypeStatus(RfqPkTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProductUse RFQ產品用途更新
        [HttpPost]
        public void UpdateProductUse(int ProductUseId = -1, string ProductUseNo = "", string ProductUseName = "", string TypeOne = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProductUseManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProductUse(ProductUseId, ProductUseNo, ProductUseName, TypeOne, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProductUseStatus -- RFQ產品用途狀態更新

        [HttpPost]
        public void UpdateProductUseStatus(int ProductUseId = -1)
        {
            try
            {
                WebApiLoginCheck("ProductUseManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProductUseStatus(ProductUseId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateMember 電商註冊資訊更新
        [HttpPost]
        public void UpdateMember(int MemberId = -1, string OrgShortName = "", string MemberName = "", string MemberEmail = "", string ContactName = ""
            , string ContactPhone = "", string ContactEmail = "", string Description = "", string Address = "", int OrgId = -1, int OrganizaitonType = -1
            , int OrganizaitonTypeId = -1, string OrganizaitonScale = "", string OrganizationCode = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("EcRegisterManagment", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateMember(MemberId, OrgShortName, MemberName, MemberEmail, ContactName
                    , ContactPhone, ContactEmail, Description, Address, OrgId, OrganizaitonType
                    , OrganizaitonTypeId, OrganizaitonScale, OrganizationCode, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMemberStatus -- 電商註冊狀態更新

        [HttpPost]
        public void UpdateMemberStatus(int MemberId = -1)
        {
            try
            {
                WebApiLoginCheck("EcRegisterManagment", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateMemberStatus(MemberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdatePerformanceGoals 業務績效目標更新
        [HttpPost]
        public void UpdatePerformanceGoals(int PgId = -1, string PgName = "", string PgDesc = "", string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePerformanceGoals(PgId, PgName, PgDesc, StartDate, EndDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePerformanceGoalsStatus 業務績效目標狀態更新
        [HttpPost]
        public void UpdatePerformanceGoalsStatus(int PgId = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePerformanceGoalsStatus(PgId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePgDetail 業務績效目標單身更新
        [HttpPost]
        public void UpdatePgDetail(int PgDetailId = -1, int PgId = -1, string IntentType = "", int UserId = -1, int IntentAmount = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePgDetail(PgDetailId, PgId, IntentType, UserId, IntentAmount);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePgDetailConfirm 業務績效目標單身確認
        [HttpPost]
        public void UpdatePgDetailConfirm(int PgDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "confirm");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePgDetailConfirm(PgDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePgDetailReConfirm 業務績效目標單身確認
        [HttpPost]
        public void UpdatePgDetailReConfirm(int PgDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "reconfirm");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdatePgDetailReConfirm(PgDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProductTypeGroup 新增產品群組資料
        [HttpPost]
        public void UpdateProductTypeGroup(int ProTypeGroupId = -1, string TypeOne = "", int RfqProClassId = -1, string ProTypeGroupName = "", string CoatingFlag = "")
        {
            try
            {
                WebApiLoginCheck("ProductTypeGroupManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProductTypeGroup(ProTypeGroupId, TypeOne, RfqProClassId, ProTypeGroupName, CoatingFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //UpdateProductTypeGroupStatus 更新產品群組狀態資料
        [HttpPost]
        public void UpdateProductTypeGroupStatus(int ProTypeGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("ProductTypeGroupManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProductTypeGroupStatus(ProTypeGroupId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //UpdateProductType 更新產品類別資料
        [HttpPost]
        public void UpdateProductType(int ProTypeId, int ProTypeGroupId = -1, string ProTypeName = "")
        {
            try
            {
                WebApiLoginCheck("ProductTypeManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProductType(ProTypeId, ProTypeGroupId, ProTypeName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //UpdateTemplateRfiSignFlowSort 更新市場評估單簽核流程排序
        [HttpPost]
        public void UpdateTemplateRfiSignFlowSort(int ProTypeGroupId = -1, int DepartmentId = -1, string SortList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateTemplateRfiSignFlowSort(ProTypeGroupId, DepartmentId, SortList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateDesignSignFlowSort 更新設計申請單流程簽核流程排序
        [HttpPost]
        public void UpdateTemplateDesignSignFlowSort(int ProTypeGroupId = -1, int DepartmentId = -1, string SortList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateTemplateDesignSignFlowSort(ProTypeGroupId, DepartmentId, SortList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProductTypeStatus 更新產品類別資料狀態
        [HttpPost]
        public void UpdateProductTypeStatus(int ProTypeId)
        {
            try
            {
                WebApiLoginCheck("ProductTypeManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProductTypeStatus(ProTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //UpdateQuotationTag 更新報價單屬性標籤
        [HttpPost]
        public void UpdateQuotationTag(int QtId = -1, string TagNo = "", string TagName = "")
        {
            try
            {
                WebApiLoginCheck("QuotationTagManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateQuotationTag(QtId, TagNo, TagName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationTagStatus 更新報價單屬性標籤狀態
        [HttpPost]
        public void UpdateQuotationTagStatus(int QtId = -1)
        {
            try
            {
                WebApiLoginCheck("QuotationTagManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateQuotationTagStatus(QtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSupplierMachineStatus 更新供應商機台資訊狀態
        [HttpPost]
        public void UpdateSupplierMachineStatus(int SmId = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "status-switch");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateSupplierMachineStatus(SmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSupplierProcessEquipment 更新參與托外掃碼的供應商的製程與機台清單
        [HttpPost]
        public void UpdateSupplierProcessEquipment(int SmId = -1, int SupplierId = -1, int ProcessId = -1, int MachineId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateSupplierProcessEquipment(SmId, SupplierId, ProcessId, MachineId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProject 更新專案資料
        [HttpPost]
        public void UpdateProject(int ProjectId = -1, string ProjectNo = "", string ProjectName = "", string EffectiveDate = "", string ExpirationDate = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProject(ProjectId, ProjectNo, ProjectName, EffectiveDate, ExpirationDate, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectDetail 更新專案預算資料
        [HttpPost]
        public void UpdateProjectDetail(int ProjectDetailId = -1, string ProjectType = "", string Currency = "", double ExchangeRate = -1, double BudgetAmount = -1, double LocalBudgetAmount = -1, string Remark = "", string ProjectFile = "")
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProjectDetail(ProjectDetailId, ProjectType, Currency, ExchangeRate, BudgetAmount, LocalBudgetAmount, Remark, ProjectFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CancelChangeProject 取消專案預算變更
        [HttpPost]
        public void CancelChangeProject(int ProjectDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.CancelChangeProject(ProjectDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CloseProject 專案結案
        [HttpPost]
        public void CloseProject(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.CloseProject(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectManualSynchronize 專案資料手動同步
        [HttpPost]
        public void UpdateProjectManualSynchronize()
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "sync");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateProjectManualSynchronize();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeletePackingVolume 包材體積資料刪除
        [HttpPost]
        public void DeletePackingVolume(int VolumeId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingVolume", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeletePackingVolume(VolumeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePackingWeight 包材體積資料刪除
        [HttpPost]
        public void DeletePackingWeight(int WeightId = -1)
        {
            try
            {
                WebApiLoginCheck("PackingWeight", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeletePackingWeight(WeightId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteItemWeight 物件重量資料刪除
        [HttpPost]
        public void DeleteItemWeight(int ItemDefaultWeightId = -1)
        {
            try
            {
                WebApiLoginCheck("ItemDefaultWeight", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteItemWeight(ItemDefaultWeightId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteForwarder 貨運承攬商資料刪除
        [HttpPost]
        public void DeleteForwarder(int ForwarderId = -1)
        {
            try
            {
                WebApiLoginCheck("ForwarderManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteForwarder(ForwarderId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDeliveryCustomer 送貨客戶資料刪除
        [HttpPost]
        public void DeleteDeliveryCustomer(int DcId = -1)
        {
            try
            {
                WebApiLoginCheck("DeliveryCustomer", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteDeliveryCustomer(DcId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePacking 包材資料刪除
        [HttpPost]
        public void DeletePacking(int PackingId = -1)
        {
            try
            {
                WebApiLoginCheck("Packing", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeletePacking(PackingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRfqProductClass -- RFQ產品類型刪除 -- Chia Yuan 230630
        [HttpPost, ValidateAntiForgeryToken]
        public void DeleteRfqProductClass(int RfqProClassId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteRfqProductClass(RfqProClassId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRfqProductType -- RFQ產品類別刪除 -- Chia Yuan 230703
        [HttpPost, ValidateAntiForgeryToken]
        public void DeleteRfqProductType(int RfqProTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "rfqproduct-type");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteRfqProductType(RfqProTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRfqPackageType -- RFQ包裝種類刪除 -- Chia Yuan 230703
        [HttpPost, ValidateAntiForgeryToken]
        public void DeleteRfqPackageType(int RfqPkTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqProductClassManagement", "rfqproduct-package");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteRfqPackageType(RfqPkTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProductUse RFQ產品用途資料刪除
        [HttpPost]
        public void DeleteProductUse(int ProductUseId = -1)
        {
            try
            {
                WebApiLoginCheck("ProductUseManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteProductUse(ProductUseId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePerformanceGoals 業務績效目標刪除
        [HttpPost]
        public void DeletePerformanceGoals(int PgId = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeletePerformanceGoals(PgId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePgDetail 業務績效目標刪除
        [HttpPost]
        public void DeletePgDetail(int PgDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("PerformanceGoalsManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeletePgDetail(PgDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProductTypeGroup 刪除產品群組資料
        [HttpPost]
        public void DeleteProductTypeGroup(int ProTypeGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("ProductTypeGroupManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteProductTypeGroup(ProTypeGroupId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //DeleteTemplateRfiSignFlow 刪除產品類別資料狀態
        [HttpPost]
        public void DeleteTemplateRfiSignFlow(int TempRfiSfId)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteTemplateRfiSignFlow(TempRfiSfId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //DeleteTemplateDesignSignFlow 刪除設計申請單流程簽核流程
        [HttpPost]
        public void DeleteTemplateDesignSignFlow(int TempDesignSfId)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteTemplateDesignSignFlow(TempDesignSfId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //DeleteProductType 刪除產品類別資料狀態
        [HttpPost]
        public void DeleteProductType(int ProTypeId)
        {
            try
            {
                WebApiLoginCheck("ProductTypeManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteProductType(ProTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //DeleteQuotationTag 刪除報價單屬性標籤狀態
        [HttpPost]
        public void DeleteQuotationTag(int QtId = -1)
        {
            try
            {
                WebApiLoginCheck("QuotationTagManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteQuotationTag(QtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteSupplierProcessEquipment 刪除參與托外掃碼的供應商製程與機台清單
        [HttpPost]
        public void DeleteSupplierProcessEquipment(int SmId = -1)
        {
            try
            {
                WebApiLoginCheck("SupplierManagement", "delete");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.DeleteSupplierProcessEquipment(SmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateCustomerSynchronize 客戶資料同步
        [HttpPost]
        [Route("api/ERP/CustomerSynchronize")]
        public void UpdateCustomerSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateCustomerSynchronize");
                #endregion

                #region //Request
                dataRequest = scmBasicInformationDA.UpdateCustomerSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateInventorySynchronize 庫別資料同步
        [HttpPost]
        [Route("api/ERP/InventorySynchronize")]
        public void UpdateInventorySynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateInventorySynchronize");
                #endregion

                #region //Request
                dataRequest = scmBasicInformationDA.UpdateInventorySynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSupplierSynchronize 供應商資料同步
        [HttpPost]
        [Route("api/ERP/SupplierSynchronize")]
        public void UpdateSupplierSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateSupplierSynchronize");
                #endregion

                #region //Request
                dataRequest = scmBasicInformationDA.UpdateSupplierSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateExchangeRateSynchronize 匯率資料同步
        [HttpPost]
        [Route("api/ERP/ExchangeRateSynchronize")]
        public void UpdateExchangeRateSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateExchangeRateSynchronize");
                #endregion

                #region //Request
                dataRequest = scmBasicInformationDA.UpdateExchangeRateSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSupplierInfo 供應商資料同步
        [HttpPost]
        [Route("api/ERP/GetSupplierInfo")]
        public void GetSupplierInfo(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateSupplierSynchronize");
                #endregion

                #region //Request
                dataRequest = scmBasicInformationDA.UpdateSupplierSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetItemInventory 品號庫存資料同步
        [HttpPost]
        [Route("api/ERP/GetItemInventory")]
        public void GetItemInventory(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateItemInventorySynchronize");
                #endregion

                #region //Request
                dataRequest = scmBasicInformationDA.UpdateItemInventorySynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //BPM相關
        #region //TransferProjectToBpm 拋轉專案預算資料到BPM
        [HttpPost]
        public void TransferProjectToBpm(int ProjectDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectBudgetManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.TransferProjectToBpm(ProjectDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectBpmRequest 接收專案預算BPM簽核狀態
        [HttpPost]
        public void GetProjectBpmRequest(string id = "", string status = "", string rootId = "", string comfirmUser = "", string bpmNo = "")
        {
            try
            {
                var dataRequestJson = new JObject();

                #region //紀錄LOG
                dataRequest = scmBasicInformationDA.AddProjectLog(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser, "P");
                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //更改MES專案預算狀態
                dataRequest = scmBasicInformationDA.UpdateProjectStatus(Convert.ToInt32(id), status, comfirmUser);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetChangeProjectBpmRequest 接收專案預算變更 BPM簽核狀態
        [HttpPost]
        public void GetChangeProjectBpmRequest(string id = "", string status = "", string rootId = "", string comfirmUser = "", string bpmNo = "")
        {
            try
            {
                var dataRequestJson = new JObject();

                #region //紀錄LOG
                dataRequest = scmBasicInformationDA.AddProjectLog(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser, "C");
                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //更改MES專案預算狀態
                dataRequest = scmBasicInformationDA.UpdateChangeProjectStatus(Convert.ToInt32(id), status, comfirmUser);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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