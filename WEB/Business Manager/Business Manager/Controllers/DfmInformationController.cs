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
    public class DfmInformationController : WebController
    {
        private DfmInformationDA dfmInformationDA = new DfmInformationDA();
        // GET: DesignForManufacturing

        public ActionResult DfmMaterialManagment()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult DesignForManufacturingManagement()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult DesignForManufacturingManagementPE()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult DfmQiProcessManagement()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult DfmMaterialManagement()
        {
            ViewLoginCheck();
            return View();
        }


        #region//Get
        #region//GetDfmItemUserAuthority --DFM使用者權限獲取
        [HttpPost]
        public void GetDfmItemUserAuthority(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmItemUserAuthority(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmItemSimple --取得品DFM項目(下拉用)
        [HttpPost]
        public void GetDfmItemSimple(int DfmItemId =-1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmItemSimple(DfmItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmTemplateParameter --取得DFM欄位數值
        [HttpPost]
        public void GetDfmTemplateParameter(int RfqProTypeId=-1,int ProductUseId=-1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read,constrained-data");

                #region //Request 
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmTemplateParameter(RfqProTypeId, ProductUseId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDesignForManufacturing --取得DFM 單頭
        [HttpPost]
        public void GetDesignForManufacturing(int DfmId = -1, string DfmNo = "", string RfqNo = "", string DifficultyLevel = "", int RfqDetailId =-1
            , int DfmItemCategoryId = -1, int RdUserId = -1, string StartCreateDate = "", string EndCreateDate = "", string Status = "", string RFQDocType = "", string AssemblyName =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request 
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDesignForManufacturing(DfmId, DfmNo, RfqNo, DifficultyLevel, RfqDetailId
                    , DfmItemCategoryId, RdUserId, StartCreateDate, EndCreateDate, Status, RFQDocType, AssemblyName
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

        #region//GetDfmDetail --取得DFM 單身
        public void GetDfmDetail(int DfmId =-1, string Version = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmDetail(DfmId, Version);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDetailMaxNum --取得DFM 單身
        public void GetDetailMaxNum(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDetailMaxNum(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmVersion --取得品DFM版本
        [HttpPost]
        public void GetDfmVersion(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmVersion(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmQuotationItem --取得DFM所需報價項目
        public void GetDfmQuotationItem(int DfmId = -1, int DfmQiId = -1, string DfmNo = "", string MaterialStatus = "", string DfmQiProcessStatus = ""
            , string DataType = "", int LoginUserId = -1,string RFQDocType =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmQuotationItem(DfmId, DfmQiId, DfmNo, MaterialStatus, DfmQiProcessStatus
                    , DataType, LoginUserId, RFQDocType
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

        #region//GetDfmQuotationItemTree --取得DFM所需報價項目(樹狀)
        public void GetDfmQuotationItemTree(int DfmId = -1, int OpenedLevel = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmQuotationItemTree(DfmId, OpenedLevel);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmQiMaterial --取得DFM使用物料資料
        public void GetDfmQiMaterial(int DfmQiId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmQiMaterial(DfmQiId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmQiOSP --取得DFM使用委外加工
        public void GetDfmQiOSP(int DfmQiId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmQiOSP(DfmQiId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmQiProcess --取得DFM製程成本
        public void GetDfmQiProcess(int DfmQiId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmQiProcess(DfmQiId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetDfmQiProcessMaxNum --取得DFM製程成本最大號
        public void GetDfmQiProcessMaxNum(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmQiProcessMaxNum(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDesignForManufacturingForSpreadsheer 取得DFM資料(Spreadsheet用)
        [HttpPost]
        public void GetDesignForManufacturingForSpreadsheer(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDesignForManufacturingForSpreadsheer(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDfmItem 取得DFM項目資料
        [HttpPost]
        public void GetDfmItem(int DfmItemId = -1, string DfmItemNo = "", string DfmItemName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "read");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.GetDfmItem(DfmItemId, DfmItemNo, DfmItemName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region//AddDfmHead --建立DFM單頭
        [HttpPost]
        public void AddDfmHead(int RfqDetailId =-1, string DfmDate = "", string DifficultyLevel ="",int RdUserId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.AddDfmHead(RfqDetailId, DfmDate, DifficultyLevel, RdUserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmDetail --建立DFM單身
        [HttpPost]
        public void AddDfmDetail(int DfmId =-1, string DfmContentList = "", string VersionControl = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.AddDfmDetail(DfmId, DfmContentList, VersionControl);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmQuotationItem --建立DFM報價項目 
        [HttpPost]
        public void AddDfmQuotationItem(int DfmId = -1, string DfmQuotationName ="", int DfmItemCategoryId = -1, int ParentDfmQiId = -1, string MfgFlag = "",
            string MaterialStatus = "", string DfmQiProcessStatus = "", double StandardCost = -1, string QuoteModel = "", double QuoteNum = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.AddDfmQuotationItem(DfmId, DfmQuotationName, DfmItemCategoryId, ParentDfmQiId, MfgFlag,
                    MaterialStatus, DfmQiProcessStatus, StandardCost, QuoteModel, QuoteNum);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmQiMaterial --建立DFM使用物料
        [HttpPost]
        public void AddDfmQiMaterial(int DfmQiId = -1, string DfmqmContentList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.AddDfmQiMaterial(DfmQiId, DfmqmContentList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmQiOSP --建立DFM使用委外加工
        [HttpPost]
        public void AddDfmQiOSP(int DfmQiId = -1, string DfmqiospContentList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.AddDfmQiOSP(DfmQiId, DfmqiospContentList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddDfmQiProcess --建立DFM使用委外加工
        [HttpPost]
        public void AddDfmQiProcess(int DfmQiId = -1, string DfmqpContentList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.AddDfmQiProcess(DfmQiId, DfmqpContentList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//UpdateDfmHead --更新DFM單頭
        [HttpPost]
        public void UpdateDfmHead(int DfmId = -1, string DfmDate = "", string DifficultyLevel ="", int RdUserId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "update");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UpdateDfmHead(DfmId, DfmDate, DifficultyLevel, RdUserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmHeadStatus --更新DFM單頭狀態
        [HttpPost]
        public void UpdateDfmHeadStatus(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "status-switch");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UpdateDfmHeadStatus(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmQuotationItem --更新DFM報價項目 
        [HttpPost]
        public void UpdateDfmQuotationItem(int DfmQiId =-1, int DfmId = -1, string DfmQuotationName = "", int DfmItemCategoryId = -1, string MfgFlag = "",
            string MaterialStatus = "", string DfmQiProcessStatus = "", double StandardCost = -1, string QuoteModel = "", double QuoteNum = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UpdateDfmQuotationItem(DfmQiId, DfmId, DfmQuotationName, DfmItemCategoryId, MfgFlag,
                    MaterialStatus, DfmQiProcessStatus, StandardCost, QuoteModel, QuoteNum);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//updateConfirmDfm --更新DFM報價項目 
        [HttpPost]
        public void updateConfirmDfm(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.updateConfirmDfm(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//updateConfirmDfmDataStatus --報價項目所需維護資料確認 
        [HttpPost]
        public void updateConfirmDfmDataStatus(int DfmQiId = -1, string Data = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.updateConfirmDfmDataStatus(DfmQiId, Data);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDfmQuotationItemLevel 更新DFM報價項目階層
        [HttpPost]
        public void UpdateDfmQuotationItemLevel(int DfmQiId = -1, int QuotationLevel = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "update");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UpdateDfmQuotationItemLevel(DfmQiId, QuotationLevel);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateDfmQiExcelData --暫存DfmQuotationItemExcel資料 
        [HttpPost]
        public void UpdateDfmQiExcelData(int DfmQiId = -1, string DataType = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UpdateDfmQiExcelData(DfmQiId, DataType, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//updateConfirmDfmQiExcelData --報價項目所需維護資料確認(Excel)
        [HttpPost]
        public void updateConfirmDfmQiExcelData(int DfmQiId = -1, string DataType = "", string SpreadsheetJson = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.updateConfirmDfmQiExcelData(DfmQiId, DataType, SpreadsheetJson, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//updateReConfirmDfmQiExcelData --報價項目所需維護資料取消確認(Excel)
        [HttpPost]
        public void updateReConfirmDfmQiExcelData(int DfmQiId = -1, string DataType = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "add");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.updateReConfirmDfmQiExcelData(DfmQiId, DataType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//DeleteDfmDetail --刪除DFM單身
        [HttpPost]
        public void DeleteDfmDetail(string DfmDetailIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmDetail(DfmDetailIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmQuotationItem --刪除DFM所需報價項目
        [HttpPost]
        public void DeleteDfmQuotationItem(int DfmQiId = -1, int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmQuotationItem(DfmQiId, DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmQiMaterial --刪除DFM使用物料
        [HttpPost]
        public void DeleteDfmQiMaterial(string DfmQiMaterialIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmQiMaterial(DfmQiMaterialIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmQiOSP --刪除DFM使用委外加工
        [HttpPost]
        public void DeleteDfmQiOSP(string DfmQiOSPIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmQiOSP(DfmQiOSPIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteDfmQiProcess --刪除DFM製程成本管理
        [HttpPost]
        public void DeleteDfmQiProcess(string DfmQiProcessIdList = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmQiProcess(DfmQiProcessIdList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDfmDetailTempSpreadsheet -- 刪除Dfm詳細資料暫存記錄
        [HttpPost]
        public void DeleteDfmDetailTempSpreadsheet(int DfmId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmDetailTempSpreadsheet(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDfmQiExcelData -- 刪除DfmQiExcelData暫存記錄
        [HttpPost]
        public void DeleteDfmQiExcelData(int DfmQiId = -1, string DataType = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "delete");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.DeleteDfmQiExcelData(DfmQiId, DataType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Uploda
        #region //UploadDfmDetail 解析SpreadSheet及上傳Dfm詳細資料
        [HttpPost]
        public void UploadDfmDetail(int DfmId, string SpreadsheetJson = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "update");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UploadDfmDetail(DfmId, SpreadsheetJson, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UploadDfmDetailTempSpreadsheetJson 暫存Dfm詳細資料SpreadsheetData
        [HttpPost]
        public void UploadDfmDetailTempSpreadsheetJson(int DfmId = -1, string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("DesignForManufacturingManagement", "update");

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.UploadDfmDetailTempSpreadsheetJson(DfmId, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //TxTransation
        #region//TxQuotationCalculation --計算DFM報價
        [HttpPost]
        public void TxQuotationCalculation(int DfmId = -1)
        {
            try
            {
                //WebApiLoginCheck();

                #region //Request
                dfmInformationDA = new DfmInformationDA();
                dataRequest = dfmInformationDA.TxQuotationCalculation(DfmId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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