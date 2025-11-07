using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NiceLabel.SDK;
using System.IO;
using System.Text;
using System.Reflection;
using MESDA;
using ClosedXML.Excel;

namespace Business_Manager.Controllers
{
    public class ToolController : WebController
    {
        private ToolDA toolDA = new ToolDA();

        // GET: Tool
        private ILabel label;
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ToolGroupManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ToolClassManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ToolCategoryManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ToolModelManagment()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ToolManagment()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ToolSpecManagment()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ToolInventoryManagment()
        {
            ViewLoginCheck();

            return View();
        }
        #region//Get

        #region //GetToolGroup 取得工具群組
        [HttpPost]
        public void GetToolGroup(int ToolGroupId = -1, string ToolGroupNo = "", string ToolGroupName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolGroupManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolGroup(ToolGroupId, ToolGroupNo, ToolGroupName, Status
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

        #region //GetToolClass 取得工具類別
        [HttpPost]
        public void GetToolClass(int ToolClassId = -1, int ToolGroupId = -1, string ToolClassNo = "", string ToolClassName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolClassManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolClass(ToolClassId, ToolGroupId, ToolClassNo, ToolClassName, Status
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

        #region //GetToolCategory 取得工具種類
        [HttpPost]
        public void GetToolCategory(int ToolCategoryId = -1, int ToolClassId = -1 , int ToolGroupId = -1, string ToolCategoryNo = "", string ToolCategoryName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolCategoryManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolCategory(ToolCategoryId, ToolClassId , ToolGroupId, ToolCategoryNo, ToolCategoryName, Status
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

        #region //GetToolModel 取得工具型號
        [HttpPost]
        public void GetToolModel(int ToolModelId = -1, int ToolCategoryId = -1, int ToolClassId = -1, int ToolGroupId = -1, int SupplierId = -1, string ToolModelNo = "", string ToolModelName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolModel(ToolModelId, ToolCategoryId, ToolClassId, ToolGroupId, SupplierId, ToolModelNo, ToolModelName, Status
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

        #region //GetTool 取得工具明細
        [HttpPost]
        public void GetTool(int ToolId = -1, int ToolModelId = -1, int ToolCategoryId = -1, int ToolClassId = -1, int ToolGroupId = -1, int SupplierId = -1, string ToolNo = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetTool(ToolId, ToolModelId, ToolCategoryId, ToolClassId, ToolGroupId, SupplierId, ToolNo, Status
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

        #region //GetToolNoMaxsor 取得工具目前最大編號
        [HttpPost]
        public void GetToolNoMaxsor(int ToolId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolNoMaxsor(ToolId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetToolSpec 取得工具規格
        [HttpPost]
        public void GetToolSpec(int ToolSpecId = -1, string ToolSpecNo = "", string ToolSpecName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolSpecManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolSpec(ToolSpecId, ToolSpecNo, ToolSpecName, Status
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

        #region //GetToolModelSpec 取得工具型號規格
        [HttpPost]
        public void GetToolModelSpec(int ToolModelSpecId = -1, int ToolModelId = -1, int ToolSpecId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "detail,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolModelSpec(ToolModelSpecId, ToolModelId, ToolSpecId, Status
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

        #region //GetToolSpecLog 取得工具型號規格
        [HttpPost]
        public void GetToolSpecLog(int ToolSpecLogId = -1, int ToolId = -1, int ToolSpecId = -1, string ToolSpecNo = "", string ToolSpecName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "detail,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolSpecLog(ToolSpecLogId, ToolId, ToolSpecId, ToolSpecNo, ToolSpecName, Status
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

        #region //GetToolInventory 取得工具倉庫
        [HttpPost]
        public void GetToolInventory(int ToolInventoryId = -1, string ToolInventoryNo = "", string ToolInventoryName = "", int ShopId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolInventory(ToolInventoryId, ToolInventoryNo, ToolInventoryName, ShopId, Status
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

        #region //GetToolLocator 取得工具儲位
        [HttpPost]
        public void GetToolLocator(int ToolLocatorId = -1, int ToolInventoryId = -1, string ToolLocatorNo = "", string ToolLocatorName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "detail,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolLocator(ToolLocatorId, ToolInventoryId, ToolLocatorNo, ToolLocatorName, Status
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

        #region //GetToolTrade 取得工具交易資料
        [HttpPost]
        public void GetToolTrade(int ToolId = -1, int ToolModelId = -1, int ToolCategoryId = -1, int ToolClassId = -1, int ToolGroupId = -1, int SupplierId = -1, string ToolNo = ""
            , string Status = "", string Source = "",string ToolKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolTrade(ToolId, ToolModelId, ToolCategoryId, ToolClassId, ToolGroupId, SupplierId, ToolNo
                    , Status, Source, ToolKey
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

        #region //GetToolTransactions 取得工具儲位
        [HttpPost]
        public void GetToolTransactions(int ToolTransactionsId = -1, int ToolId = -1, string TransactionType = "", string StartDate = "", string EndDate = "", int TraderId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "detail,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolTransactions(ToolTransactionsId, ToolId, TransactionType, StartDate, EndDate, TraderId
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

        #region //GetLabelPrintMachine 取得標籤機資訊
        [HttpPost]
        public void GetLabelPrintMachine(int LabelPrintId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ToolGroupManagment", "read,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetLabelPrintMachine(LabelPrintId
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

        #region //GetToolCountInLocator --取得該儲位入庫數
        [HttpPost]
        public void GetToolCountInLocator(int ToolLocatorId = -1)
        {
            try
            {
                // WebApiLoginCheck("ToolInventoryManagment", "detail,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetToolCountInLocator(ToolLocatorId);

                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLocatorTool --取得該儲位入庫刀具明細
        [HttpPost]
        public void GetLocatorTool(int ToolLocatorId = -1)
        {
            try
            {
                // WebApiLoginCheck("ToolInventoryManagment", "detail,constrained-data");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.GetLocatorTool(ToolLocatorId);

                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//Add

        #region //AddToolSpec 工具規格新增
        [HttpPost]
        public void AddToolSpec(string ToolSpecNo = "", string ToolSpecName = "")
        {
            try
            {
                WebApiLoginCheck("ToolSpecManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolSpec(ToolSpecNo, ToolSpecName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolGroup 工具群組新增
        [HttpPost]
        public void AddToolGroup(string ToolGroupNo = "", string ToolGroupName = "")
        {
            try
            {
                WebApiLoginCheck("ToolGroupManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolGroup(ToolGroupNo, ToolGroupName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolClass 工具類別新增
        [HttpPost]
        public void AddToolClass(int ToolGroupId = -1, string ToolClassNo = "", string ToolClassName = "")
        {
            try
            {
                WebApiLoginCheck("ToolClassManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolClass(ToolGroupId, ToolClassNo, ToolClassName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolCategory 工具種類新增
        [HttpPost]
        public void AddToolCategory(int ToolClassId = -1, string ToolCategoryNo = "", string ToolCategoryName = "")
        {
            try
            {
                WebApiLoginCheck("ToolCategoryManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolCategory(ToolClassId, ToolCategoryNo, ToolCategoryName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolModel 工具型號新增
        [HttpPost]
        public void AddToolModel(int ToolCategoryId = -1, string ToolModelNo = "", string ToolModelErpNo = "", string ToolModelName = "", int SupplierId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolModel(ToolCategoryId, ToolModelNo, ToolModelErpNo, ToolModelName, SupplierId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolModelSpec 工具型號規格新增
        [HttpPost]
        public void AddToolModelSpec(int ToolModelId = -1, int ToolSpecId = -1, Double ToolSpecValue = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolModelSpec(ToolModelId, ToolSpecId, ToolSpecValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTool 工具明細資料新增
        [HttpPost]
        public void AddTool(int ToolModelId = -1, string ToolNo = "")
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddTool(ToolModelId, ToolNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolBatch 工具明細資料新增(批量)
        [HttpPost]
        public void AddToolBatch(int ToolModelId = -1, int ProductionNum = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolBatch(ToolModelId, ProductionNum);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolSpecLog 工具明細規格新增
        [HttpPost]
        public void AddToolSpecLog(int ToolId = -1, int ToolSpecId = -1, Double ToolSpecValue = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolSpecLog(ToolId, ToolSpecId, ToolSpecValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolInventory 工具倉庫資料新增
        [HttpPost]
        public void AddToolInventory(string ToolInventoryNo = "", string ToolInventoryName = "", string ToolInventoryDesc = "", int ShopId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolInventory(ToolInventoryNo, ToolInventoryName, ToolInventoryDesc, ShopId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolLocator 工具儲位資料新增
        [HttpPost]
        public void AddToolLocator(string ToolLocatorNo = "", string ToolLocatorName = "", string ToolLocatorDesc = "", int ToolInventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolLocator(ToolLocatorNo, ToolLocatorName, ToolLocatorDesc, ToolInventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddToolTransactions 工具交易入庫新增
        [HttpPost]
        public void AddToolTransactions(int ToolId = -1, string TransactionDate = "", string TransactionType = "", int ToolLocatorId = -1
            , string TransactionReason = "", int TraderId = -1, string IsLimit = "", int LastToolId = -1, int ProcessingQty = -1)
        {
            try
            {
                //WebApiLoginCheck("ProductionHistory", "add");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.AddToolTransactions(ToolId, TransactionDate, TransactionType, ToolLocatorId
                    , TransactionReason, TraderId, IsLimit, LastToolId, ProcessingQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //UpdateToolGroup 工具群組資料更新
        [HttpPost]
        public void UpdateToolGroup(int ToolGroupId = -1, string ToolGroupName = "")
        {
            try
            {
                WebApiLoginCheck("ToolGroupManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolGroup(ToolGroupId, ToolGroupName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolGroupStatus 工具群組狀態更新
        [HttpPost]
        public void UpdateToolGroupStatus(int ToolGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolGroupManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolGroupStatus(ToolGroupId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolClass 工具類別資料更新
        [HttpPost]
        public void UpdateToolClass(int ToolClassId = -1, int ToolGroupId = -1, string ToolClassName = "")
        {
            try
            {
                WebApiLoginCheck("ToolClassManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolClass(ToolClassId, ToolGroupId, ToolClassName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolClassStatus 工具類別狀態更新
        [HttpPost]
        public void UpdateToolClassStatus(int ToolClassId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolClassManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolClassStatus(ToolClassId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolCategory 工具種類資料更新
        [HttpPost]
        public void UpdateToolCategory(int ToolCategoryId = -1, int ToolClassId = -1, string ToolCategoryName = "")
        {
            try
            {
                WebApiLoginCheck("ToolCategoryManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolCategory(ToolCategoryId, ToolClassId, ToolCategoryName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolCategoryStatus 工具種類狀態更新
        [HttpPost]
        public void UpdateToolCategoryStatus(int ToolCategoryId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolCategoryManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolCategoryStatus(ToolCategoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolModel 工具型號資料更新
        [HttpPost]
        public void UpdateToolModel(int ToolModelId = -1, int ToolCategoryId = -1, string ToolModelErpNo = "", string ToolModelName = "", int SupplierId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolModel(ToolModelId, ToolCategoryId, ToolModelErpNo, ToolModelName, SupplierId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolModelStatus 工具型號狀態更新
        [HttpPost]
        public void UpdateToolModelStatus(int ToolModelId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolModelStatus(ToolModelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTool 工具明細資料更新
        [HttpPost]
        public void UpdateTool(int ToolId = -1, int ToolModelId = -1, int RevToolCategoryId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateTool(ToolId, ToolModelId, RevToolCategoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolStatus 工具型號狀態更新
        [HttpPost]
        public void UpdateToolStatus(int ToolId = -1,string ToolNo = "")
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolStatus(ToolId, ToolNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolSpec 工具規格資料更新
        [HttpPost]
        public void UpdateToolSpec(int ToolSpecId = -1, string ToolSpecName = "")
        {
            try
            {
                WebApiLoginCheck("ToolSpecManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolSpec(ToolSpecId, ToolSpecName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolSpecStatus 工具規格狀態更新
        [HttpPost]
        public void UpdateToolSpecStatus(int ToolSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolSpecManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolSpecStatus(ToolSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolModelSpec 工具型號規格資料更新
        [HttpPost]
        public void UpdateToolModelSpec(int ToolModelSpecId = -1, int ToolSpecId = -1, Double ToolSpecValue = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolModelSpec(ToolModelSpecId, ToolSpecId, ToolSpecValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolModelSpecStatus 工具規格狀態更新
        [HttpPost]
        public void UpdateToolModelSpecStatus(int ToolModelSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolModelSpecStatus(ToolModelSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolSpecLog 工具型號規格資料更新
        [HttpPost]
        public void UpdateToolSpecLog(int ToolSpecLogId = -1, int ToolSpecId = -1, Double ToolSpecValue = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolSpecLog(ToolSpecLogId, ToolSpecId, ToolSpecValue);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolSpecLogStatus 工具規格狀態更新
        [HttpPost]
        public void UpdateToolSpecLogStatus(int ToolSpecLogId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolSpecLogStatus(ToolSpecLogId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolInventory 工具倉庫資料更新
        [HttpPost]
        public void UpdateToolInventory(int ToolInventoryId = -1, string ToolInventoryName = "", string ToolInventoryDesc = "", int ShopId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "update");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolInventory(ToolInventoryId, ToolInventoryName, ToolInventoryDesc, ShopId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolInventoryStatus 工具倉庫狀態更新
        [HttpPost]
        public void UpdateToolInventoryStatus(int ToolInventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "status-switch");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolInventoryStatus(ToolInventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolLocator 工具倉庫資料更新
        [HttpPost]
        public void UpdateToolLocator(int ToolLocatorId = -1, string ToolLocatorName = "", string ToolLocatorDesc = "")
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolLocator(ToolLocatorId, ToolLocatorName, ToolLocatorDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateToolLocatorStatus 工具倉庫狀態更新
        [HttpPost]
        public void UpdateToolLocatorStatus(int ToolLocatorId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.UpdateToolLocatorStatus(ToolLocatorId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//Delete

        #region //DeleteToolGroup 工具群組刪除
        [HttpPost]
        public void DeleteToolGroup(int ToolGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolGroupManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolGroup(ToolGroupId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolClass 工具類別刪除
        [HttpPost]
        public void DeleteToolClass(int ToolClassId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolClassManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolClass(ToolClassId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolCategory 工具種類刪除
        [HttpPost]
        public void DeleteToolCategory(int ToolCategoryId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolCategoryManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolCategory(ToolCategoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolModel 工具型號刪除
        [HttpPost]
        public void DeleteToolModel(int ToolModelId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolModel(ToolModelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTool 工具明細刪除
        [HttpPost]
        public void DeleteTool(int ToolId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteTool(ToolId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolSpec 工具規格刪除
        [HttpPost]
        public void DeleteToolSpec(int ToolSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolSpecManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolSpec(ToolSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolModelSpec 工具型號規格刪除
        [HttpPost]
        public void DeleteToolModelSpec(int ToolModelSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolModelManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolModelSpec(ToolModelSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolSpecLog 工具明細規格刪除
        [HttpPost]
        public void DeleteToolSpecLog(int ToolSpecLogId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolSpecLog(ToolSpecLogId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolInventory 工具倉庫刪除
        [HttpPost]
        public void DeleteToolInventory(int ToolInventoryId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "delete");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolInventory(ToolInventoryId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteToolLocator 工具儲位刪除
        [HttpPost]
        public void DeleteToolLocator(int ToolLocatorId = -1)
        {
            try
            {
                WebApiLoginCheck("ToolInventoryManagment", "detail");

                #region //Request
                toolDA = new ToolDA();
                dataRequest = toolDA.DeleteToolLocator(ToolLocatorId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//Print

        #region //PrintToolList -- 工具列印(直接)  -- Shintokuro 2023-01-06
        [HttpPost]
        public void PrintToolList(string ToolList, string PrintMachine ,string PrintType)
        {
            try
            {
                toolDA = new ToolDA();

                string LabelPath = "", line2;
                //PrintMachine = "TSC TDP-345-1";
                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\LabelName.txt", Encoding.Default);
                if (PrintType == "1D")
                {
                    LabelPathTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\LabelName.txt", Encoding.Default);
                }
                else if(PrintType == "QR")
                {
                    LabelPathTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\LabelNameQR.txt", Encoding.Default);

                }
                
                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "cab MACH/300-01";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                else
                {
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印

                string ToolNo = "";
                var ToolArr = ToolList.Split(',');

                for (var i = 0; i < ToolArr.Length; i++)
                {
                    ToolNo = ToolArr[i];
                    dataRequest = toolDA.GetTool(-1, -1, -1, -1, -1, -1, ToolArr[i], ""
                    , "", -1, -1);
                    #region //Response
                    string ToolModelErpNo = "";
                    string ToolCategoryName = "";
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        ToolModelErpNo = result[0]["ToolModelErpNo"].ToString();
                        ToolCategoryName = result[0]["ToolCategoryName"].ToString();
                    }
                    #endregion

                    label.Variables["KNIFE_NO"].SetValue(ToolArr[i]);
                    label.Variables["KNIFE_TYPE"].SetValue(ToolModelErpNo);
                    label.Variables["KNIFE_CATEGORY"].SetValue(ToolCategoryName);
                    label.Print(1);
                    dataRequest = toolDA.UpdateToolStatus(-1, ToolArr[i]);

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion
                    if (jsonResponse.First.First.ToString() == "errorForDA")
                    {
                        break;
                    }
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
        #endregion

        #region //PrintToolTemporary -- 工具暫存列表列印  -- Shintokuro 2023-01-06
        [HttpPost]
        public void PrintToolTemporary(string ToolTemporaryList , string PrintMachine, string PrintType)
        {
            try
            {
                toolDA = new ToolDA();

                string LabelPath = "", line2;
                //PrintMachine = "TSC TDP-345-1";

                #region //NiceLabel套件
                string sdkFilesPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles");
                if (Directory.Exists(sdkFilesPath))
                {
                    PrintEngineFactory.SDKFilesPath = sdkFilesPath;
                }
                PrintEngineFactory.PrintEngine.Initialize();
                #endregion

                #region //指定要列印的標籤機與標籤檔案路徑

                #region //標籤格式圖檔路徑
                StreamReader LabelPathTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\LabelName.txt", Encoding.Default);
                if (PrintType == "1D")
                {
                    LabelPathTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\LabelName.txt", Encoding.Default);
                }
                else if (PrintType == "QR")
                {
                    LabelPathTxt = new StreamReader("C:\\WIN\\PrintToolBarcode\\Template\\LabelNameQR.txt", Encoding.Default);

                }

                while ((line2 = LabelPathTxt.ReadLine()) != null)
                {
                    LabelPath = line2.ToString();
                }
                #endregion

                #region //印製標籤機路徑
                if (PrintMachine == "")
                {
                    StreamReader PrintMachineTxt = new StreamReader("C:\\WIN\\PrintCartonBarcode\\Template\\PrinterName.txt", Encoding.Default);
                    while ((line2 = PrintMachineTxt.ReadLine()) != null)
                    {
                        PrintMachine = line2.ToString();
                        //PrintMachine = "cab MACH/300-01";
                    }
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                else
                {
                    label = PrintEngineFactory.PrintEngine.OpenLabel(LabelPath);
                    label.PrintSettings.PrinterName = PrintMachine;
                }
                #endregion

                #region //標籤列印

                string ToolNo = "";
                var ToolArr = ToolTemporaryList.Split(',');

                for (var i = 0; i < ToolArr.Length; i++)
                {
                    ToolNo = ToolArr[i];
                    dataRequest = toolDA.GetTool(-1, -1, -1, -1, -1, -1, ToolArr[i], ""
                    , "", -1, -1);
                    #region //Response
                    string ToolModelErpNo = "";
                    string ToolCategoryName = "";
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        ToolModelErpNo = result[0]["ToolModelErpNo"].ToString();
                        ToolCategoryName = result[0]["ToolCategoryName"].ToString();
                    }
                    #endregion

                    label.Variables["KNIFE_NO"].SetValue(ToolArr[i]);
                    label.Variables["KNIFE_TYPE"].SetValue(ToolModelErpNo);
                    label.Variables["KNIFE_CATEGORY"].SetValue(ToolCategoryName);

                    label.Print(1);
                    dataRequest = toolDA.UpdateToolStatus(-1, ToolArr[i]);

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    if(jsonResponse.First.First.ToString() == "errorForDA")
                    {
                        break;
                    }
                    ;
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
        #endregion

        #endregion
    }
}