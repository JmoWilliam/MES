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
    public class InventoryTransactionController : WebController
    {
        private InventoryTransactionDA inventoryTransactionDA = new InventoryTransactionDA();

        #region //View
        public ActionResult Login()
        {
            return View();
        }

        public ActionResult Index()
        {
            InventoryTransactionLoginCheck();

            return View();
        }

        public ActionResult TransferDocument()
        {
            InventoryTransactionLoginCheck();

            return View();
        }

        public ActionResult InventoryTransactionNote()
        {
            InventoryTransactionLoginCheck();

            return View();
        }

        public ActionResult MaterialRequisition()
        {
            InventoryTransactionLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetInventoryTransaction 取得庫存異動單據資料
        [HttpPost]
        public void GetInventoryTransaction(int ItId = -1, string ItErpPrefix = "", string ItErpNo = "", string ItErpFullNo = "", string ItDate = ""
            , string DocDate = "", string MtlItemNo = "", string MtlItemName = "", int InventoryId = -1, int ToInventoryId = -1, string ConfirmStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetInventoryTransaction(ItId, ItErpPrefix, ItErpNo, ItErpFullNo, ItDate
                    , DocDate, MtlItemNo, MtlItemName, InventoryId, ToInventoryId, ConfirmStatus
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

        #region //GetItDetail 取得庫存異動單據詳細資料
        [HttpPost]
        public void GetItDetail(int ItDetailId = -1, int ItId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetItDetail(ItDetailId, ItId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlItem 取得品號資料
        [HttpPost]
        public void GetMtlItem(int MtlItemId = -1, string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetMtlItem(MtlItemId, MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetInventoryQty 取得庫存數量
        [HttpPost]
        public void GetInventoryQty(int MtlItemId = -1, int InventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetInventoryQty(MtlItemId, InventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMaterialRequisition 取得領退料單據資料
        [HttpPost]
        public void GetMaterialRequisition(int MrId = -1, string MrErpPrefix = "", string MrErpNo = "", string MrErpFullNo = "", string MrDate = ""
            , string DocDate = "", string MtlItemNo = "", string MtlItemName = "", string WoErpFullNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetMaterialRequisition(MrId, MrErpPrefix, MrErpNo, MrErpFullNo, MrDate
                    , DocDate, MtlItemNo, MtlItemName, WoErpFullNo
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

        #region //GetMrDetail 取得領退料單據詳細資料
        [HttpPost]
        public void GetMrDetail(int MrDetailId = -1, int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetMrDetail(MrDetailId, MrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetManufactureOrder 取得製令相關資料
        [HttpPost]
        public void GetManufactureOrder(string WoErpFullNo = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetManufactureOrder(WoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMrWipOrder 取得領料單製令綁定相關資料
        [HttpPost]
        public void GetMrWipOrder(int MrId = -1, int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetMrWipOrder(MrId, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMrBarcodeRegister 取得領料單條碼註冊資料
        [HttpPost]
        public void GetMrBarcodeRegister(int BarcodeRegisterId = -1, int MrDetailId = -1, string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetMrBarcodeRegister(BarcodeRegisterId, MrDetailId, BarcodeNo
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

        #region //GetMrBarcodeReRegister 取得領料單條碼退料資料
        [HttpPost]
        public void GetMrBarcodeReRegister(int BarcodeReRegisterId = -1, int MrDetailId = -1, string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetMrBarcodeReRegister(BarcodeReRegisterId, MrDetailId, BarcodeNo
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

        #region //GetStorageLocation 取得庫別儲位相關資料
        [HttpPost]
        public void GetStorageLocation(int InventoryId = -1, string StorageLocation = "", string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetStorageLocation(InventoryId, StorageLocation, MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetInventoryLocationFlag 取得庫別儲位設定
        [HttpPost]
        public void GetInventoryLocationFlag(int InventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetInventoryLocationFlag(InventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLocationQty 取得庫別儲位庫存數量
        [HttpPost]
        public void GetLocationQty(int MtlItemId = -1, int InventoryId = -1, string Location = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetLocationQty(MtlItemId, InventoryId, Location);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetAuthority 取得庫存異動相關權限資料
        [HttpPost]
        public void GetAuthority()
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetAuthority();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDeviceInventory 取得裝置庫別資料
        [HttpPost]
        public void GetDeviceInventory(string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetDeviceInventory(MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLoginUserInfo 取得使用者資料
        [HttpPost]
        public void GetLoginUserInfo()
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.GetLoginUserInfo();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddInventoryTransaction 新增庫存異動單據
        [HttpPost]
        public void AddInventoryTransaction(string ItErpPrefix = "", string ItErpNo = "", string ItDate = "", string DocDate = "", int DepartmentId = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "add");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.AddInventoryTransaction(ItErpPrefix, ItErpNo, ItDate, DocDate, DepartmentId, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddItDetail 新增庫存異動單據詳細資料
        [HttpPost]
        public void AddItDetail(int ItId = -1, int MtlItemId = -1, string ItMtlItemName = "", string ItMtlItemSpec = ""
            , double ItQty = -1, int UomId = -1, int InventoryId = -1, int ToInventoryId = -1, string ItRemark = "", string StorageLocation = "", string ToStorageLocation = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "add");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.AddItDetail(ItId, MtlItemId, ItMtlItemName, ItMtlItemSpec
                    , ItQty, UomId, InventoryId, ToInventoryId, ItRemark, StorageLocation, ToStorageLocation);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMaterialRequisition 新增領退料單據資料
        [HttpPost]
        public void AddMaterialRequisition(string MrErpPrefix = "", string MrDate = "", string DocDate = "", string ProductionLine = ""
            , string ContractManufacturer = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "add");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.AddMaterialRequisition(MrErpPrefix, MrDate, DocDate, ProductionLine
                    , ContractManufacturer, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMrWipOrder 新增領退料單據 製令帶出資料
        [HttpPost]
        public void AddMrWipOrder(int MrId = -1, int MoId = -1, string WoErpPrefix = "", string WoErpNo = "", string MrType = "", double Quantity = -1
            , string RequisitionCode = "", string NegativeStatus = "", string Remark = "", string MaterialCategory = "", string SubinventoryType = ""
            , string LineSeq = "", int InventoryId = -1, string StorageLocation = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "add");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.AddMrWipOrder(MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                    , RequisitionCode, NegativeStatus, Remark, MaterialCategory, SubinventoryType
                    , LineSeq, InventoryId, StorageLocation);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMrBarcodeRegister 新增領料單條碼註冊資料
        [HttpPost]
        public void AddMrBarcodeRegister(int MrDetailId = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "add");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.AddMrBarcodeRegister(MrDetailId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMrBarcodeReRegister 新增領料單條碼退料資料
        [HttpPost]
        public void AddMrBarcodeReRegister(int MrDetailId = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "add");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.AddMrBarcodeReRegister(MrDetailId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateInventoryTransaction 更新庫存異動單據
        [HttpPost]
        public void UpdateInventoryTransaction(int ItId = -1, string ItDate = "", string DocDate = "", int DepartmentId = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.UpdateInventoryTransaction(ItId, ItDate, DocDate, DepartmentId, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateItDetail 更新庫存異動單據詳細資料
        [HttpPost]
        public void UpdateItDetail(int ItDetailId = -1, int ItId = -1, int MtlItemId = -1, string ItMtlItemName = "", string ItMtlItemSpec = ""
            , double ItQty = -1, int UomId = -1, int InventoryId = -1, int ToInventoryId = -1, string ItRemark = "", string StorageLocation = "", string ToStorageLocation = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.UpdateItDetail(ItDetailId, ItId, MtlItemId, ItMtlItemName, ItMtlItemSpec
                    , ItQty, UomId, InventoryId, ToInventoryId, ItRemark, StorageLocation, ToStorageLocation);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMaterialRequisition 更新領退料單據
        [HttpPost]
        public void UpdateMaterialRequisition(int MrId = -1, string MrErpPrefix = "", string MrDate = "", string DocDate = "", string ProductionLine = ""
            , string ContractManufacturer = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.UpdateMaterialRequisition(MrId, MrErpPrefix, MrDate, DocDate, ProductionLine
                    , ContractManufacturer, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMrWipOrder 更新製令帶出資料設定(重計單身)
        [HttpPost]
        public void UpdateMrWipOrder(int MrId, int MoId, string WoErpPrefix, string WoErpNo, string MrType, double Quantity
            , string RequisitionCode, string NegativeStatus, string Remark, string MaterialCategory, string SubinventoryType
            , string LineSeq, int InventoryId, string StorageLocation)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.UpdateMrWipOrder(MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                    , RequisitionCode, NegativeStatus, Remark, MaterialCategory, SubinventoryType
                    , LineSeq, InventoryId, StorageLocation);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMrWipOrderNoRC 更新製令帶出資料設定(不重計單身)
        [HttpPost]
        public void UpdateMrWipOrderNoRC(int MrId, int MoId, string WoErpPrefix, string WoErpNo, string MrType, double Quantity
            , string RequisitionCode, string NegativeStatus, string Remark, string MaterialCategory, string SubinventoryType
            , string LineSeq, int InventoryId, string StorageLocation)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.UpdateMrWipOrderNoRC(MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                    , RequisitionCode, NegativeStatus, Remark, MaterialCategory, SubinventoryType, LineSeq
                    , InventoryId, StorageLocation);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMrDetail 更新領退料單據詳細資料
        [HttpPost]
        public void UpdateMrDetail(int MrDetailId = -1, int MoId = -1, int MtlItemId = -1, double Quantity = -1, int UomId = -1, int InventoryId = -1, string Remark = "", string DetailDesc = "", string StorageLocation = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "update");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.UpdateMrDetail(MrDetailId, MoId, MtlItemId, Quantity, UomId, InventoryId, Remark, DetailDesc, StorageLocation);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteInventoryTransaction -- 刪除庫存異動單據
        [HttpPost]
        public void DeleteInventoryTransaction(int ItId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "delete");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteInventoryTransaction(ItId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteItDetail -- 刪除庫存異動單據詳細資料
        [HttpPost]
        public void DeleteItDetail(int ItDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "delete");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteItDetail(ItDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMaterialRequisition -- 刪除領退料單據
        [HttpPost]
        public void DeleteMaterialRequisition(int MrId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "delete");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteMaterialRequisition(MrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMrWipOrder -- 刪除製令帶出資料
        [HttpPost]
        public void DeleteMrWipOrder(int MrId = -1, int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "delete");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteMrWipOrder(MrId, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMrDetail -- 刪除領退料單據詳細資料
        [HttpPost]
        public void DeleteMrDetail(int MrDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "delete");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteMrDetail(MrDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMrBarcodeRegister -- 刪除領料單條碼註冊資料
        [HttpPost]
        public void DeleteMrBarcodeRegister(int BarcodeRegisterId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteMrBarcodeRegister(BarcodeRegisterId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMrBarcodeReRegister -- 刪除領料單條碼退料資料
        [HttpPost]
        public void DeleteMrBarcodeReRegister(int BarcodeReRegisterId = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialRequisition", "barcode-control");

                #region //Request
                inventoryTransactionDA = new InventoryTransactionDA();
                dataRequest = inventoryTransactionDA.DeleteMrBarcodeReRegister(BarcodeReRegisterId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void InventoryTransactionLogin(string SystemKey, string Account)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetSubSystemLogin(SystemKey, Account, "InventoryTransaction");
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    LoginLog(Convert.ToInt32(result[0]["UserId"]), Account, "InventoryTransaction");
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
        public void InventoryTransactionLoginCheck()
        {
            bool verify = LoginVerify("InventoryTransaction");

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "InventoryTransaction");

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
        public void InventoryTransactionApiLoginCheck()
        {
            bool verify = LoginVerify("InventoryTransaction");

            if (!verify)
            {
                throw new SystemException("系統已登出，請重新登入");
            }
        }
        #endregion
        #endregion
    }
}