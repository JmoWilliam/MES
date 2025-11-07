using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Helpers;
using EIPDA;
using SCMDA;
using ZXing;
using ZXing.Common;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using ClosedXML.Excel;

namespace Business_Manager.Controllers
{
    public class RequestForInformationController : WebController
    {
        private RequestForQuotationDA requestForQuotationDA = new RequestForQuotationDA();
        private RequestForInformationDA requestForInformationDA = new RequestForInformationDA();
        private ScmBasicInformationDA scmBasicInformationDA = new ScmBasicInformationDA();

        #region //View
        public ActionResult ProdTerminalManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProdSystemManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProdModuleManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult ProductUse()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult TemplateRfiManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult TemplateDesignManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult RfiManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult DesignManagement()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetProdTerminal -- 取得終端資料
        [HttpPost]
        public void GetProdTerminal(int ProdTerminalId = -1, string TerminalName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProdTerminalManagement", "read,constrained-data");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetProdTerminal(ProdTerminalId, TerminalName, Status
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

        #region //GetProdSystem -- 取得系統資料
        [HttpPost]
        public void GetProdSystem(int ProdSysId = -1, string SystemName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProdSystemManagement", "read,constrained-data");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetProdSystem(ProdSysId, SystemName, Status
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

        #region //GetProdModule -- 取得模組資料
        [HttpPost]
        public void GetProdModule(int ProdModId = -1, string ModuleName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModuleManagement", "read,constrained-data");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetProdModule(ProdModId, ModuleName, Status
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

        #region //GetProductUse -- 取得產品應用資料
        [HttpPost]
        public void GetProductUse(int ProductUseId = -1, string ProductUseName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProductUseManagement", "read,constrained-data");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetProductUse(ProductUseId, ProductUseName, Status
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

        #region //GetSales -- 取得負責業務資料
        [HttpPost]
        public void GetSales(int UserId = -1, string UserName = "", int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetSales(UserId, UserName, CompanyId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTemplateProdSpec -- 取得評估樣板主資料
        [HttpPost]
        public void GetTemplateProdSpec(int TempProdSpecId = -1, int ProTypeGroupId = -1, string SpecName = "", string SpecEName = "", string ControlType = "", string Status = "", string FeatureName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "read,constrained-data");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetTemplateProdSpec(TempProdSpecId, ProTypeGroupId, SpecName, SpecEName, ControlType, Status, FeatureName
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

        #region //GetTemplateSpecParameter -- 取得新設計項目主資料
        [HttpPost]
        public void GetTemplateSpecParameter(int TempSpId = -1, int ProTypeGroupId = -1, string ParameterName = "", string ControlType = "", string Status = "", string PmtDetailName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "read,constrained-data");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetTemplateSpecParameter(TempSpId, ProTypeGroupId, ParameterName, ControlType, Status, PmtDetailName
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

        #region //GetType -- 取得類別資料
        [HttpPost]
        public void GetType(string TypeSchema = "", string TypeNo = "", string TypeName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetType(TypeSchema, TypeNo, TypeName
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

        #region //GetRequestForInformation -- 取得RFI單頭
        [HttpPost]
        public void GetRequestForInformation(int RfiId = -1, string RfiNo = "", int ProTypeGroupId = -1, int ProductUseId = -1, int SalesId = -1, string StartDate = "", string EndDate = ""
            , string Status = "", string RfiDetailStatus = "", string SignStatus = "", string FlowStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "read");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetRequestForInformation(RfiId, RfiNo, ProTypeGroupId, ProductUseId, SalesId, StartDate, EndDate
                    , Status, RfiDetailStatus, SignStatus, FlowStatus
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

        #region //GetRfiDetail -- 取得RFI單身
        [HttpPost]
        public void GetRfiDetail(int RfiId = -1, int RfiDetailId = -1, int SalesId = -1, string RfiDetailStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "read");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetRfiDetail(RfiId, RfiDetailId, SalesId, RfiDetailStatus
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

        #region //GetRfiSignFlow -- 取得RFI簽核流程
        [HttpPost]
        public void GetRfiSignFlow(int RfiSfId = -1, int RfiDetailId = -1, string Edition = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "read");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetRfiSignFlow(RfiSfId, RfiDetailId, Edition, Status
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

        #region //GetRfiDetailByTemplate -- 取得RFI單身(樣板單筆查詢) (停用)
        //[HttpPost]
        //public void GetRfiDetailByTemplate(int RfiId = -1, int RfiDetailId = -1)
        //{
        //    try
        //    {
        //        WebApiLoginCheck("RfiManagement", "read");

        //        #region //Request
        //        requestForInformationDA = new RequestForInformationDA();
        //        dataRequest = requestForInformationDA.GetRfiDetailByTemplate(RfiId, RfiDetailId);
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

        #region //GetExchangeRate -- 取得匯率
        [HttpPost]
        public void GetExchangeRate(string Currency)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetExchangeRate(Currency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetExchangeRateERP -- 取得四捨五入位數資料
        [HttpPost]
        public void GetExchangeRateERP(string Currency)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForInformationDA.GetExchangeRateERP(Currency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDemandDesign -- 取得新設計申請單
        [HttpPost]
        public void GetDemandDesign(int DesignId = -1, int RfiId = -1, string DesignNo = "", int SalesId = -1, int ProTypeGroupId = -1, int ProductUseId = -1
            , string CustomerName = "", string AssemblyName = "", string StartDate = "", string EndDate = ""
            , string SignStatus = "", string DesignStatus = "", string FlowStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForInformationDA.GetDemandDesign(DesignId, RfiId, DesignNo, SalesId, ProTypeGroupId, ProductUseId
                    , CustomerName, AssemblyName, StartDate, EndDate
                    , SignStatus, DesignStatus, FlowStatus
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

        #region //GetRfiSignFlow -- 取得設計申請簽核流程
        [HttpPost]
        public void GetDesignSignFlow(int DesignSfId = -1, int DesignId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "read");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetDesignSignFlow(DesignSfId, DesignId, Status
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

        #region //GetDemandDesignSpec -- 取得新設計規格資料
        [HttpPost]
        public void GetDemandDesignSpec(int DdSpecId = -1, int DesignId = -1, string ParameterName = "", string PmtDetailName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForInformationDA.GetDemandDesignSpec(DdSpecId, DesignId, ParameterName, PmtDetailName, Status
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

        #region //GetDesignInput -- 取得設計規格特徵輸入值資料
        [HttpPost]
        public void GetDesignInput(int DdsDetailId = -1, int DdSpecId = -1, string PmtDetailName = "", string DesignInput = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForInformationDA.GetDesignInput(DdsDetailId, DdSpecId, PmtDetailName, DesignInput, Status
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

        #region //GetDesignHistory -- 取得設計規格特徵歷史紀錄
        [HttpPost]
        public void GetDesignHistory(int DesignId = -1, string SchemaSpecEName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForInformationDA.GetDesignHistory(DesignId, SchemaSpecEName
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

        #region //CheckSignAuthority -- 權限檢核 From DetailCode
        private void CheckSignAuthority(string FunctionCode = "", string DetailCode = "", string Status = "")
        {
            requestForInformationDA = new RequestForInformationDA();
            dataRequest = requestForInformationDA.GetAuthority(FunctionCode, DetailCode, Status
                , "", -1, -1);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    throw new SystemException("無此單據簽核權限，不得查看!");
                }
            }
            else
            {
                throw new SystemException(jsonResponse["msg"].ToString());
            }
        }
        #endregion

        #region //CheckRfiAdditionalFile -- 附件資料檢核
        private void CheckRfiAdditionalFile(int RfiId = -1, int FileId = -1)
        {
            requestForInformationDA = new RequestForInformationDA();
            dataRequest = requestForInformationDA.GetRequestForInformation(RfiId, "", -1, -1, -1, "", "", "", "", "", ""
                , "", -1, -1);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                    throw new SystemException("【市場評估單】資料錯誤!");

                foreach (var item in jsonResponse["result"])
                {
                    if (!item["RfiDetail"].ToString().TryParseJson(out JObject rdiDetail))
                        throw new SystemException("【市場評估單】單身資料錯誤!");

                    var detail = rdiDetail["data"].Where(w => !string.IsNullOrWhiteSpace(w["AdditionalFile"]?.ToString()));
                    if (!detail.Any())
                        throw new SystemException("【市場評估單】檔案錯誤!");
                    else if (!detail.Any(a => Convert.ToInt32(a["AdditionalFile"]?.ToString()) == FileId))
                        throw new SystemException("【市場評估單】檔案錯誤!");
                }
            }
            else
            {
                throw new SystemException(jsonResponse["msg"].ToString());
            }
        }
        #endregion

        #region //DownloadDesignFile -- 下載檔案-New
        public ActionResult DownloadDesignFile(int FileId = -1, int DesignId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "read,constrained-data");

                #region //檢查簽核權限
                //CheckSignAuthority("DesignManagement", "SignRD,SignMP,SignSS,SignSRD", "A");
                if (BaseHelper.ClientLinkType() != "廠內連線") throw new SystemException("非廠內IP，不得下載附件!");
                #endregion

                #region //檢查單據是否錯誤
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetDesignSignAuthority(DesignId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() != "success")
                    throw new SystemException(jsonResponse["msg"].ToString());

                return RedirectToAction("GetFile", "Web", new { fileId = FileId });
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
                Response.ContentType = "application/json";
                Response.Write(jsonResponse.ToString(Formatting.None));

                return new EmptyResult();
            }
        }
        #endregion

        #region //GetRfiSignFlowFile -- 下載檔案(無來源驗證)
        public ActionResult GetRfiSignFlowFile(int FileId = -1, int RfiSfId = -1)
        {
            try
            {
                #region //檢查簽核權限
                if (BaseHelper.ClientLinkType() != "廠內連線") throw new SystemException("非廠內IP，不得下載附件!");
                #endregion

                #region //檢查單據是否錯誤
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetRfiSignFlow(RfiSfId, -1, "", ""
                    , "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() != "success")
                    throw new SystemException(jsonResponse["msg"].ToString());

                return RedirectToAction("GetFile", "Web", new { fileId = FileId });
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
                Response.ContentType = "application/json";
                Response.Write(jsonResponse.ToString(Formatting.None));

                return new EmptyResult();
            }
        }
        #endregion

        #region //GetDesignSignFlowFile -- 下載檔案(無來源驗證)
        public ActionResult GetDesignSignFlowFile(int FileId = -1, int DesignSfId = -1)
        {
            try
            {
                #region //檢查簽核權限
                if (BaseHelper.ClientLinkType() != "廠內連線") throw new SystemException("非廠內IP，不得下載附件!");
                #endregion

                #region //檢查單據是否錯誤
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetDesignSignFlow(DesignSfId, -1, ""
                    , "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() != "success")
                    throw new SystemException(jsonResponse["msg"].ToString());

                return RedirectToAction("GetFile", "Web", new { fileId = FileId });
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
                Response.ContentType = "application/json";
                Response.Write(jsonResponse.ToString(Formatting.None));

                return new EmptyResult();
            }
        }
        #endregion
        #endregion

        #region //Add
        #region //AddProdTerminal 終端資料新增
        [HttpPost]
        public void AddProdTerminal(string TerminalName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProdTerminalManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddProdTerminal(TerminalName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProdSystem 系統資料新增
        [HttpPost]
        public void AddProdSystem(string SystemName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProdSystemManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddProdSystem(SystemName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProdModule 系統資料新增
        [HttpPost]
        public void AddProdModule(string ModuleName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProdModuleManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddProdModule(ModuleName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateProdSpec 評估樣板主資料新增
        [HttpPost]
        public void AddTemplateProdSpec(int ProTypeGroupId = -1, string SpecName = "", string SpecEName = "", string ControlType = "", string Required = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddTemplateProdSpec(ProTypeGroupId, SpecName, SpecEName, ControlType, Required, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateSpecParameter 設計樣板主資料新增
        [HttpPost]
        public void AddTemplateSpecParameter(int ProTypeGroupId = -1, string ParameterName = "", string ParameterEName = "", string ControlType = "", string Required = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddTemplateSpecParameter(ProTypeGroupId, ParameterName, ParameterEName, ControlType, Required, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateProdSpecDetail 評估樣板子資料新增
        [HttpPost]
        public void AddTemplateProdSpecDetail(int TempProdSpecId = -1, string FeatureName = "", string DataType = "", string Description = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddTemplateProdSpecDetail(TempProdSpecId, FeatureName, DataType, Description, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateSpecParameterDetail 設計樣板子資料新增
        [HttpPost]
        public void AddTemplateSpecParameterDetail(int TempSpId = -1, string PmtDetailName = "", string DataType = "", string Description = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddTemplateSpecParameterDetail(TempSpId, PmtDetailName, DataType, Description, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateRfiSignFlow 市場評估單簽核流程新增
        [HttpPost]
        public void AddTemplateRfiSignFlow(int ProTypeGroupId = -1, int DepartmentId = -1, string UserList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddTemplateRfiSignFlow(ProTypeGroupId, DepartmentId, UserList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CopyTemplateRfiSignFlow 複製市場評估單流程
        [HttpPost]
        public void CopyTemplateRfiSignFlow(int ProTypeGroupId = -1, int BaseDepartmentId = -1, int DepartmentId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.CopyTemplateRfiSignFlow(ProTypeGroupId, BaseDepartmentId, DepartmentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateDesignSignFlow 設計申請單簽核流程新增
        [HttpPost]
        public void AddTemplateDesignSignFlow(int ProTypeGroupId = -1, int DepartmentId = -1, string UserList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.AddTemplateDesignSignFlow(ProTypeGroupId, DepartmentId, UserList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //CopyTemplateDesignSignFlow 複製新設計申請單流程
        [HttpPost]
        public void CopyTemplateDesignSignFlow(int ProTypeGroupId = -1, int BaseDepartmentId = -1, int DepartmentId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "add");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.CopyTemplateDesignSignFlow(ProTypeGroupId, BaseDepartmentId, DepartmentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRequestForInformation RFI單頭資料新增
        [HttpPost]
        public void AddRequestForInformation(string RfiNo = "", int ProTypeGroupId = -1, int ProductUseId = -1, int SalesId = -1, string ExistFlag = ""
            , string CustomerName = "", string CustomerEName = "", int CustomerId = -1, string OrganizaitonType = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddRequestForInformation(RfiNo, ProTypeGroupId, ProductUseId, SalesId, ExistFlag
                    , CustomerName, CustomerEName, CustomerId, OrganizaitonType, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRfiDetail RFI單身新增
        [HttpPost]
        public void AddRfiDetail(int RfiId = -1, string RfiSequence = "", int SalesId = -1, string ProdDate = "", string ProdLifeCycleStart = "", string ProdLifeCycleEnd = ""
            , decimal TargetUnitPrice = -1, int LifeCycleQty = -1, decimal Revenue = -1, decimal RevenueFC = -1
            , string Currency = "", decimal ExchangeRate = -1, int AdditionalFile = -1
            , string ProdSpecs = "", string ProdSpecDetails = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddRfiDetail(RfiId, RfiSequence, SalesId, ProdDate, ProdLifeCycleStart, ProdLifeCycleEnd
                    , TargetUnitPrice, LifeCycleQty, Revenue, RevenueFC
                    , Currency, ExchangeRate, AdditionalFile
                    , ProdSpecs, ProdSpecDetails);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDemandDesignSpec 新設計申請單規格新增
        [HttpPost]
        public void AddDemandDesignSpec(int DesignId = -1, string ProdSpecs = "", string ProdSpecDetails = "")
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "add");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.AddDemandDesignSpec(DesignId, ProdSpecs, ProdSpecDetails);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateProdTerminal 終端資料更新
        [HttpPost]
        public void UpdateProdTerminal(int ProdTerminalId = -1, string TerminalName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProdTerminalManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProdTerminal(ProdTerminalId, TerminalName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdTerminalStatus 終端資料狀態更新
        [HttpPost]
        public void UpdateProdTerminalStatus(int ProdTerminalId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdTerminalManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProdTerminalStatus(ProdTerminalId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdSystem 系統資料更新
        [HttpPost]
        public void UpdateProdSystem(int ProdSysId = -1, string SystemName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProdSystemManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProdSystem(ProdSysId, SystemName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdSystemStatus 系統資料狀態更新
        [HttpPost]
        public void UpdateProdSystemStatus(int ProdSysId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdSystemManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProdSystemStatus(ProdSysId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdModule 模組資料更新
        [HttpPost]
        public void UpdateProdModule(int ProdModId = -1, string ModuleName = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("ProdModuleManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProdModule(ProdModId, ModuleName, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdModuleStatus 模組資料狀態更新
        [HttpPost]
        public void UpdateProdModuleStatus(int ProdModId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModuleManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProdModuleStatus(ProdModId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProductTypeGroupStatus 產品群組狀態更新
        [HttpPost]
        public void UpdateProductTypeGroupStatus(int ProTypeGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProductTypeGroupStatus(ProTypeGroupId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProductTypeStatus 產品類別狀態更新
        [HttpPost]
        public void UpdateProductTypeStatus(int ProTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateProductTypeStatus(ProTypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpec 評估樣板主資料更新
        [HttpPost]
        public void UpdateTemplateProdSpec(int TempProdSpecId = -1, string SpecName = "", string SpecEName = "",  string Required = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpec(TempProdSpecId, SpecName, SpecEName, Required, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameter 設計樣板主資料更新
        [HttpPost]
        public void UpdateTemplateSpecParameter(int TempSpId = -1, string ParameterName = "", string ParameterEName = "", string Required = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameter(TempSpId, ParameterName, ParameterEName, Required, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpecDetail 評估樣板子資料更新
        [HttpPost]
        public void UpdateTemplateProdSpecDetail(int TempPsDetailId = -1, string FeatureName = "", string Description = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpecDetail(TempPsDetailId, FeatureName, Description, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameterDetail 設計樣板子資料更新
        [HttpPost]
        public void UpdateTemplateSpecParameterDetail(int TempSpDetailId = -1, string PmtDetailName = "", string Description = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameterDetail(TempSpDetailId, PmtDetailName, Description, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpecControlType 評估樣板控制項類型更新
        [HttpPost]
        public void UpdateTemplateProdSpecControlType(int TempProdSpecId = -1, string ControlType = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpecControlType(TempProdSpecId, ControlType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameterControlType 設計樣板控制項類型更新
        [HttpPost]
        public void UpdateTemplateSpecParameterControlType(int TempSpId = -1, string ControlType = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameterControlType(TempSpId, ControlType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateRfiSignFlow 更新設計申請單流程簽核流程排序
        [HttpPost]
        public void UpdateTemplateRfiSignFlow(int TempRfiSfId = -1, string UserList = "", string FlowJobName = "", string FlowStatus = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateTemplateRfiSignFlow(TempRfiSfId, UserList, FlowJobName, FlowStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateDesignSignFlow 更新設計申請單流程簽核流程排序
        [HttpPost]
        public void UpdateTemplateDesignSignFlow(int TempDesignSfId = -1, string UserList = "", string FlowJobName = "", string FlowStatus = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.UpdateTemplateDesignSignFlow(TempDesignSfId, UserList, FlowJobName, FlowStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpecStatus 評估樣板資料狀態更新
        [HttpPost]
        public void UpdateTemplateProdSpecStatus(int TempProdSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpecStatus(TempProdSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameterStatus 設計樣板資料狀態更新
        [HttpPost]
        public void UpdateTemplateSpecParameterStatus(int TempSpId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameterStatus(TempSpId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpecRequired 評估樣板必填注記更新
        [HttpPost]
        public void UpdateTemplateProdSpecRequired(int TempProdSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpecRequired(TempProdSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameterRequired 設計樣板必填注記更新
        [HttpPost]
        public void UpdateTemplateSpecParameterRequired(int TempSpId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameterRequired(TempSpId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpecSort 評估樣板主資料排序
        [HttpPost]
        public void UpdateTemplateProdSpecSort(int ProTypeGroupId = -1, string SortList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpecSort(ProTypeGroupId, SortList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateProdSpecDetailSort 評估樣板子資料排序
        [HttpPost]
        public void UpdateTemplateProdSpecDetailSort(int TempProdSpecId = -1, string SortList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateProdSpecDetailSort(TempProdSpecId, SortList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameterSort 設計樣板主資料排序
        [HttpPost]
        public void UpdateTemplateSpecParameterSort(int ProTypeGroupId = -1, string SortList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameterSort(ProTypeGroupId, SortList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateSpecParameterDetailSort 設計樣板子資料排序
        [HttpPost]
        public void UpdateTemplateSpecParameterDetailSort(int TempSpId = -1, string SortList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateTemplateSpecParameterDetailSort(TempSpId, SortList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRequestForInformation RFI單頭資料更新
        [HttpPost]
        public void UpdateRequestForInformation(int RfiId = -1, int ProTypeGroupId = -1, int ProductUseId = -1, int SalesId = -1, string ExistFlag = ""
            , string CustomerName = "", string CustomerEName = "", int CustomerId = -1, string OrganizaitonType = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateRequestForInformation(RfiId, ProTypeGroupId, ProductUseId, SalesId, ExistFlag
                    , CustomerName, CustomerEName, CustomerId, OrganizaitonType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRequestForInformationStatus RFI單頭資料更新
        [HttpPost]
        public void UpdateRequestForInformationStatus(int RfiId = -1, string StatusNo = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateRequestForInformationStatus(RfiId, StatusNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfiDetail RFI單身資料更新
        [HttpPost]
        public void UpdateRfiDetail(int RfiDetailId = -1, string ProdDate = "", string ProdLifeCycleStart = "", string ProdLifeCycleEnd = ""
            , decimal TargetUnitPrice = -1, decimal LifeCycleQty = -1, decimal Revenue = -1, decimal RevenueFC = -1
            , string Currency = "", decimal ExchangeRate = -1, int AdditionalFile = -1, string ProdSpecs = "", string ProdSpecDetails = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateRfiDetail(RfiDetailId, ProdDate, ProdLifeCycleStart, ProdLifeCycleEnd
                    , TargetUnitPrice, LifeCycleQty, Revenue, RevenueFC
                    , Currency, ExchangeRate, AdditionalFile, ProdSpecs, ProdSpecDetails);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfiDetailAdditionalFile RFI單身附檔資料更新
        [HttpPost]
        public void UpdateRfiDetailAdditionalFile(int RfiDetailId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateRfiDetailAdditionalFile(RfiDetailId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDesignFile RFI單身附檔資料更新
        [HttpPost]
        public void UpdateDesignFile(int DesignId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateDesignFile(DesignId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfiDetailStatus RFI單身狀態更新
        [HttpPost]
        public void UpdateRfiDetailStatus(int RfiDetailId = -1, int RfiSfId = -1, int DepartmentId = -1, string RfiDetailStatus = "", string SignContent = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateRfiDetailStatus(RfiDetailId, RfiSfId, DepartmentId, RfiDetailStatus, SignContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //RfiUrgentSignMessage 發送RFI催簽通知信
        [HttpPost]
        public void RfiUrgentSignMessage(int RfiDetailId = -1, int RfiSfId = -1, string UrgentSignContent = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.RfiUrgentSignMessage(RfiDetailId, RfiSfId, UrgentSignContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfiSignFlowAdditionalFile RFI簽核附檔資料更新
        [HttpPost]
        public void UpdateRfiSignFlowAdditionalFile(int RfiSfId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateRfiSignFlowAdditionalFile(RfiSfId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDemandDesignStatus 新設計申請單狀態更新
        [HttpPost]
        public void UpdateDemandDesignStatus(int DesignId = -1, int DesignSfId = -1, string DesignStatus = "", string SignContent = "")
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateDemandDesignStatus(DesignId, DesignSfId, DesignStatus, SignContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DesignUrgentSignMessage 發送新設計申請催簽通知信
        [HttpPost]
        public void DesignUrgentSignMessage(int DesignSfId = -1, int DesignId = -1, string UrgentSignContent = "")
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "status-switch");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DesignUrgentSignMessage(DesignSfId, DesignId, UrgentSignContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDesignSignFlowAdditionalFile 新設計申請簽核附檔資料更新
        [HttpPost]
        public void UpdateDesignSignFlowAdditionalFile(int DesignSfId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateDesignSignFlowAdditionalFile(DesignSfId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDemandDesignSpec 新設計申請單規格更新
        [HttpPost]
        public void UpdateDemandDesignSpec(int DesignId = -1, string ProdSpecs = "", string ProdSpecDetails = "")
        {
            try
            {
                WebApiLoginCheck("DesignManagement", "update");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.UpdateDemandDesignSpec(DesignId, ProdSpecs, ProdSpecDetails);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteProdTerminal 終端資料刪除
        [HttpPost]
        public void DeleteProdTerminal(int ProdTerminalId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdTerminalManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteProdTerminal(ProdTerminalId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProdSystem 系統資料刪除
        [HttpPost]
        public void DeleteProdSystem(int ProdSysId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdSystemManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteProdSystem(ProdSysId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProdModule 模組資料刪除
        [HttpPost]
        public void DeleteProdModule(int ProdModId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModuleManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteProdModule(ProdModId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateProdSpec 評估樣板資料刪除
        [HttpPost]
        public void DeleteTemplateProdSpec(int TempProdSpecId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteTemplateProdSpec(TempProdSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateProdSpecDetail -- 評估樣板子資料刪除
        [HttpPost]
        public void DeleteTemplateProdSpecDetail(int TempPsDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateRfiManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteTemplateProdSpecDetail(TempPsDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateSpecParameter 設計樣板資料刪除
        [HttpPost]
        public void DeleteTemplateSpecParameter(int TempSpId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteTemplateSpecParameter(TempSpId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateSpecParameterDetail -- 設計樣板子資料刪除
        [HttpPost]
        public void DeleteTemplateSpecParameterDetail(int TempSpDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateDesignManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteTemplateSpecParameterDetail(TempSpDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRequestForInformation -- RFI單頭資料刪除
        [HttpPost]
        public void DeleteRequestForInformation(int RfiId = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteRequestForInformation(RfiId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRfiDetail --RFI單身資料刪除
        [HttpPost]
        public void DeleteRfiDetail(int RfiDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "delete");

                #region //Request
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.DeleteRfiDetail(RfiDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //自定義儲存格
        #region //StatusStyle
        public IXLStyle GenCellStyle(string status)
        {
            var xlstyle = XLWorkbook.DefaultStyle;
            xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            xlstyle.Border.TopBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.RightBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.TopBorderColor = XLColor.Black;
            xlstyle.Border.BottomBorderColor = XLColor.Black;
            xlstyle.Border.LeftBorderColor = XLColor.Black;
            xlstyle.Border.RightBorderColor = XLColor.Black;
            xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
            xlstyle.Font.FontSize = 12;
            xlstyle.Font.Bold = false;
            xlstyle.Protection.SetLocked(false);
            switch (status)
            {
                case "總經理":
                    xlstyle.Fill.BackgroundColor = XLColor.FromName("CadetBlue");
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
                case "業務":
                case "研發":
                case "業務主管":
                case "研發主管":
                    xlstyle.Fill.BackgroundColor = XLColor.Yellow;
                    break;
                case "新設計申請單":
                case "回饋客戶":
                case "完成評估":
                    xlstyle.Fill.BackgroundColor = XLColor.Green;
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
                case "不准":
                case "退回":
                case "退回/變更":
                    xlstyle.Fill.BackgroundColor = XLColor.FromName("Firebrick");
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
                case "評估單核准":
                    xlstyle.Fill.BackgroundColor = XLColor.FromIndex(25);
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
                default:
                    break;
            }
            return xlstyle;
        }
        #endregion

        public IXLStyle GetCellStyle(string style)
        {
            var xlstyle = XLWorkbook.DefaultStyle;
            switch (style)
            {
                case "defaultStyle":
                    xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    xlstyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.TopBorderColor = XLColor.Black;
                    xlstyle.Border.BottomBorderColor = XLColor.Black;
                    xlstyle.Border.LeftBorderColor = XLColor.Black;
                    xlstyle.Border.RightBorderColor = XLColor.Black;
                    xlstyle.Fill.BackgroundColor = XLColor.NoColor;
                    xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    xlstyle.Font.FontSize = 12;
                    xlstyle.Font.Bold = false;
                    xlstyle.Protection.SetLocked(false);
                    break;
                case "titleStyle":
                    xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    xlstyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.TopBorderColor = XLColor.Black;
                    xlstyle.Border.BottomBorderColor = XLColor.Black;
                    xlstyle.Border.LeftBorderColor = XLColor.Black;
                    xlstyle.Border.RightBorderColor = XLColor.Black;
                    xlstyle.Fill.BackgroundColor = XLColor.NoColor;
                    xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    xlstyle.Font.FontSize = 16;
                    xlstyle.Font.Bold = false;
                    break;
                case "headerStyle":
                    xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    xlstyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    xlstyle.Border.TopBorderColor = XLColor.Black;
                    xlstyle.Border.BottomBorderColor = XLColor.Black;
                    xlstyle.Border.LeftBorderColor = XLColor.Black;
                    xlstyle.Border.RightBorderColor = XLColor.Black;
                    xlstyle.Fill.BackgroundColor = XLColor.AliceBlue;
                    xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    xlstyle.Font.FontSize = 14;
                    xlstyle.Font.Bold = true;
                    break;
                case "dateStyle":
                    xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    xlstyle.Border.TopBorder = XLBorderStyleValues.None;
                    xlstyle.Border.BottomBorder = XLBorderStyleValues.None;
                    xlstyle.Border.LeftBorder = XLBorderStyleValues.None;
                    xlstyle.Border.RightBorder = XLBorderStyleValues.None;
                    xlstyle.Border.TopBorderColor = XLColor.NoColor;
                    xlstyle.Border.BottomBorderColor = XLColor.NoColor;
                    xlstyle.Border.LeftBorderColor = XLColor.NoColor;
                    xlstyle.Border.RightBorderColor = XLColor.NoColor;
                    xlstyle.Fill.BackgroundColor = XLColor.NoColor;
                    xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    xlstyle.Font.FontSize = 12;
                    xlstyle.Font.Bold = false;
                    xlstyle.DateFormat.Format = "yyyy-mm-dd HH:mm:ss";
                    break;
                case "numberStyle":
                    xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    xlstyle.Border.TopBorder = XLBorderStyleValues.None;
                    xlstyle.Border.BottomBorder = XLBorderStyleValues.None;
                    xlstyle.Border.LeftBorder = XLBorderStyleValues.None;
                    xlstyle.Border.RightBorder = XLBorderStyleValues.None;
                    xlstyle.Border.TopBorderColor = XLColor.NoColor;
                    xlstyle.Border.BottomBorderColor = XLColor.NoColor;
                    xlstyle.Border.LeftBorderColor = XLColor.NoColor;
                    xlstyle.Border.RightBorderColor = XLColor.NoColor;
                    xlstyle.Fill.BackgroundColor = XLColor.NoColor;
                    xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    xlstyle.Font.FontSize = 12;
                    xlstyle.Font.Bold = false;
                    xlstyle.NumberFormat.Format = "_-* #,##0.00_-;-* #,##0.00_-;_-* \" - \"??_-;_-@_-";
                    break;
                case "currencyStyle":
                    xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    xlstyle.Border.TopBorder = XLBorderStyleValues.None;
                    xlstyle.Border.BottomBorder = XLBorderStyleValues.None;
                    xlstyle.Border.LeftBorder = XLBorderStyleValues.None;
                    xlstyle.Border.RightBorder = XLBorderStyleValues.None;
                    xlstyle.Border.TopBorderColor = XLColor.NoColor;
                    xlstyle.Border.BottomBorderColor = XLColor.NoColor;
                    xlstyle.Border.LeftBorderColor = XLColor.NoColor;
                    xlstyle.Border.RightBorderColor = XLColor.NoColor;
                    xlstyle.Fill.BackgroundColor = XLColor.NoColor;
                    xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    xlstyle.Font.FontSize = 12;
                    xlstyle.Font.Bold = false;
                    xlstyle.DateFormat.Format = "$#,##0.00";
                    break;
            }
            return xlstyle;
        }

        #endregion

        #region //Excel
        #region //GetRfiExcel
        public void GetRfiExcel(int RfiId = -1, string RfiNo = "", int ProTypeGroupId = -1, int ProductUseId = -1, int SalesId = -1, string StartDate = "", string EndDate = "", string SignStatus = "", string RfiDetailStatus = "", string FlowStatus = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "excel");

                #region //Request - RFI單頭資料
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetRfiToExcel(RfiId, RfiNo, ProTypeGroupId, ProductUseId, SalesId, StartDate, EndDate, SignStatus, RfiDetailStatus, FlowStatus, "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    var proTypeGroup = JObject.Parse(dataRequest)["data"].Children<JObject>().Select(s => new { ProTypeGroupName = s["產品類別"]?.ToString() }).Distinct().ToList();

                    byte[] excelFile;
                    string excelFileName = "RFI市場評估單Excel檔" + StartDate + "_" + EndDate;

                    using (var workbook = new XLWorkbook())
                    {
                        proTypeGroup.ForEach(f =>
                        {
                            string excelsheetName = f.ProTypeGroupName;

                            //int.TryParse(f.ProTypeGroupId, out int proTypeGroupId);

                            #region //Request - RFI單頭資料
                            //requestForInformationDA = new RequestForInformationDA();
                            //dataRequest = requestForInformationDA.GetRfiToExcel(RfiId, RfiNo, proTypeGroupId, ProductUseId, SalesId, StartDate, EndDate, SignStatus, RfiDetailStatus, FlowStatus, "", -1, -1);
                            #endregion

                            if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                            {
                                var data = JObject.Parse(dataRequest)["data"].Children<JObject>().Where(w => w["產品類別"]?.ToString() == f.ProTypeGroupName).ToList();

                                #region //參數初始化
                                List<string> header = data.Take(1).Properties().Select(s => s.Name).ToList();
                                string colIndex = "";
                                int rowIndex = 1;
                                #endregion

                                var worksheet = workbook.Worksheets.Add(excelsheetName);
                                worksheet.RowHeight = 16;
                                worksheet.Style = GetCellStyle("defaultStyle");

                                #region //圖片
                                var imagePath = Server.MapPath("~/Content/images/logo.png");
                                var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 7)).Scale(0.07);
                                worksheet.Row(rowIndex).Height = 40;
                                worksheet.Range(rowIndex, 1, rowIndex, header.Count).Merge().Style = GetCellStyle("titleStyle");
                                rowIndex++;
                                #endregion

                                #region //HEADER
                                for (int i = 0; i < header.Count; i++)
                                {
                                    colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                    worksheet.Cell(colIndex).SetValue(header[i]);
                                    worksheet.Cell(colIndex).Style = GetCellStyle("headerStyle");
                                }
                                #endregion

                                #region //BODY
                                foreach (var item in data)
                                {
                                    int i = 1;
                                    rowIndex++;
                                    header.ForEach(h =>
                                    {
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(i, rowIndex)).SetValue(item[h]?.ToString());
                                        if (h == "關卡狀態") worksheet.Cell(BaseHelper.MergeNumberToChar(i, rowIndex)).Style = GenCellStyle(item[h]?.ToString());
                                        i++;
                                    });
                                }
                                #endregion

                                #region //設定
                                #region //篩選
                                worksheet.RangeUsed().SetAutoFilter(); // set filter
                                #endregion

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
                            }
                        });

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

        #region //GetDesignExcel
        public void GetDesignExcel(int DesignId = -1, string DesignNo = "", int ProTypeGroupId = -1, int ProductUseId = -1, int SalesId = -1, string StartDate = "", string EndDate = "", string SignStatus = "", string DesignStatus = "", string FlowStatus = "")
        {
            try
            {
                WebApiLoginCheck("RfiManagement", "excel");

                #region //Request - RFI單頭資料
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetDesignToExcel(DesignId, DesignNo, ProTypeGroupId, ProductUseId, SalesId, StartDate, EndDate, SignStatus, DesignStatus, FlowStatus, "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    var proTypeGroup = JObject.Parse(dataRequest)["data"].Children<JObject>().Select(s => new { ProTypeGroupName = s["產品類別"]?.ToString() }).Distinct().ToList();

                    byte[] excelFile;
                    string excelFileName = "RFI新設計申請單Excel檔" + StartDate + "_" + EndDate;

                    using (var workbook = new XLWorkbook())
                    {
                        proTypeGroup.ForEach(f =>
                        {
                            string excelsheetName = f.ProTypeGroupName;

                            //int.TryParse(f["ProTypeGroupId"]?.ToString(), out int proTypeGroupId);

                            #region //Request - RFI單頭資料
                            //requestForInformationDA = new RequestForInformationDA();
                            //dataRequest = requestForInformationDA.GetDesignToExcel(DesignId, DesignNo, proTypeGroupId, ProductUseId, SalesId, StartDate, EndDate, DesignStatus, "", -1, -1);
                            #endregion

                            if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                            {
                                var data = JObject.Parse(dataRequest)["data"].Children<JObject>().Where(w => w["產品類別"]?.ToString() == f.ProTypeGroupName).ToList();

                                #region //參數初始化
                                List<string> header = data.Take(1).Properties().Select(s => s.Name).ToList();
                                string colIndex = "";
                                int rowIndex = 1;
                                #endregion

                                var worksheet = workbook.Worksheets.Add(excelsheetName);
                                worksheet.RowHeight = 16;
                                worksheet.Style = GetCellStyle("defaultStyle");

                                #region //圖片
                                var imagePath = Server.MapPath("~/Content/images/logo.png");
                                var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 7)).Scale(0.07);
                                worksheet.Row(rowIndex).Height = 40;
                                worksheet.Range(rowIndex, 1, rowIndex, header.Count).Merge().Style = GetCellStyle("titleStyle");
                                rowIndex++;
                                #endregion

                                #region //HEADER
                                for (int i = 0; i < header.Count; i++)
                                {
                                    colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                                    worksheet.Cell(colIndex).SetValue(header[i]);
                                    worksheet.Cell(colIndex).Style = GetCellStyle("headerStyle");
                                }
                                #endregion

                                #region //BODY
                                foreach (var item in data)
                                {
                                    int i = 1;
                                    rowIndex++;
                                    header.ForEach(h =>
                                    {
                                        worksheet.Cell(BaseHelper.MergeNumberToChar(i, rowIndex)).SetValue(item[h]?.ToString());
                                        if (h == "關卡狀態") worksheet.Cell(BaseHelper.MergeNumberToChar(i, rowIndex)).Style = GenCellStyle(item[h]?.ToString());
                                        i++;
                                    });
                                }
                                #endregion

                                #region //設定
                                #region //篩選
                                worksheet.RangeUsed().SetAutoFilter(); // set filter
                                #endregion

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
                            }
                        });

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

        #region //Pdf
        #region //GetRfiPdf 取得RFI Pdf
        public void GetRfiPdf(int RfiId = -1, string Status = "", string RfiDetailStatus = "", string OrderBy = "")
        {
            try
            {
                logger.Info(string.Format("市場評估單列印(開始),使用者:{0},RfiId:{1},時間:{2}", requestForInformationDA.CurrentUser, RfiId, DateTime.Now));

                WebApiLoginCheck("RfiManagement", "print");

                Regex regex_newline = new Regex("(\r\n|\r|\n)");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - RFI單頭資料
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetRequestForInformation(RfiId, "", -1, -1, -1, "", "", Status, RfiDetailStatus, "", "", OrderBy, -1, -1);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                var writer = new BarcodeWriter  //dll裡面可以看到屬性
                {
                    Format = BarcodeFormat.CODE_128,
                    Options = new EncodingOptions //設定大小
                    {
                        Height = 50,
                        Width = 200,
                        PureBarcode = false,
                        Margin = 0
                    }
                };

                string htmlText = "";

                if (jsonResponse["status"].ToString() == "success")
                {
                    //if (!jsonResponse["result"].ToString().TryParseJson(out JObject result)) throw new SystemException("【RFI單據】資料錯誤!");
                    if (jsonResponse["result"].Count() <= 0) throw new SystemException("【RFI單據】資料錯誤!");

                    #region //html
                    foreach (var rfi in jsonResponse["result"])
                    {
                        htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/RequestForInformationPdf.html"));

                        #region 產生條碼
                        var img = writer.Write(Convert.ToString(RfiId));
                        Bitmap myBitmap = new Bitmap(img);
                        string FileName = "RfiId" + rfi["RfiId"].ToString();
                        string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/RfiBarcode/{0}.bmp", FileName));
                        myBitmap.Save(filePath, ImageFormat.Bmp);
                        #endregion

                        htmlText = htmlText.Replace("[Barcode]", filePath);

                        var RfiNo = rfi["RfiNo"].ToString();
                        var AssemblyName = rfi["AssemblyName"].ToString();
                        var ProductUseName = rfi["ProductUseName"].ToString();
                        var SalesNo = rfi["SalesNo"].ToString();
                        var SalesName = rfi["SalesName"].ToString();
                        var SalesGender = rfi["SalesGender"].ToString();
                        var OrganizaitonTypeName = rfi["OrganizaitonTypeName"].ToString();
                        var CustomerName = rfi["CustomerName"].ToString();
                        var CustomerEName = rfi["CustomerEName"].ToString();
                        var RfiStatusName = rfi["RfiStatusName"].ToString();
                        var ExistFlagName = rfi["ExistFlagName"].ToString();
                        var CompanyName = rfi["CompanyName"].ToString();
                        var CreateDate = rfi["CreateDate"].ToString();

                        htmlText = htmlText.Replace("[RfiNo]", RfiNo);
                        htmlText = htmlText.Replace("[AssemblyName]", AssemblyName);
                        htmlText = htmlText.Replace("[ProductUseName]", ProductUseName);
                        htmlText = htmlText.Replace("[SalesNo]", SalesNo);
                        htmlText = htmlText.Replace("[SalesName]", SalesName);
                        htmlText = htmlText.Replace("[SalesGender]", SalesGender);
                        htmlText = htmlText.Replace("[OrganizaitonTypeName]", OrganizaitonTypeName);
                        htmlText = htmlText.Replace("[CustomerName]", CustomerName);
                        htmlText = htmlText.Replace("[CustomerEName]", CustomerEName);
                        htmlText = htmlText.Replace("[RfiStatusName]", RfiStatusName);
                        htmlText = htmlText.Replace("[ExistFlagName]", ExistFlagName);
                        htmlText = htmlText.Replace("[CompanyName]", CompanyName);
                        htmlText = htmlText.Replace("[CreateDate]", CreateDate);

                        string htmlTemplate = htmlText;
                        int pageNum = 15;

                        if (!rfi["RfiDetail"].ToString().TryParseJson(out JObject tempRfiDetail)) throw new SystemException("【市場評估單身】資料錯誤!");

                        var lastRfiDetail = tempRfiDetail["data"].OrderByDescending(o => o["Edition"]).FirstOrDefault();

                        int.TryParse(lastRfiDetail["RfiDetailId"].ToString(), out int RfiDetailId);

                        #region //Request - RFI單身資料
                        requestForInformationDA = new RequestForInformationDA();
                        string detailRequest = requestForInformationDA.GetRfiDetail(RfiId, RfiDetailId, -1, RfiDetailStatus, "", -1, -1);
                        JObject detailJsonResponse = BaseHelper.DAResponse(detailRequest);
                        #endregion

                        if (detailJsonResponse["status"].ToString() == "success")
                        {
                            if (detailJsonResponse["result"].Count() <= 0) throw new SystemException("【市場評估單身】資料錯誤!");

                            //若單身改為版本控制，則必須找出最後一版或跑迴圈
                            var rfiDetail = detailJsonResponse["result"].FirstOrDefault();

                            List<(int, string)> DictLogos = new List<(int, string)> { (2, "jmo_logo"), (3, "eterge_logo") };
                            var corp = DictLogos.Find(f => f.Item1.ToString() == rfiDetail["CompanyId"]?.ToString());
                            if (string.IsNullOrWhiteSpace(corp.Item2)) throw new SystemException("【公司別】資料錯誤!");

                            htmlText = htmlText.Replace("[CompanyLogo]", Server.MapPath(string.Format("~/Areas/SCM/Content/images/{0}.png", corp.Item2)));

                            var RfiSequence = rfiDetail["RfiSequence"].ToString();
                            var Edition = rfiDetail["Edition"].ToString();
                            decimal.TryParse(rfiDetail["TargetUnitPrice"].ToString(), out decimal TargetUnitPrice);
                            int.TryParse(rfiDetail["LifeCycleQty"].ToString(), out int LifeCycleQty);
                            decimal.TryParse(rfiDetail["Revenue"].ToString(), out decimal Revenue);
                            decimal.TryParse(rfiDetail["RevenueFC"].ToString(), out decimal RevenueFC);
                            decimal.TryParse(rfiDetail["ExchangeRate"].ToString(), out decimal ExchangeRate);
                            var CurrencyName = rfiDetail["CurrencyName"].ToString();
                            var Currency = rfiDetail["Currency"].ToString();
                            var ProdDate = rfiDetail["ProdDate"].ToString();
                            var ProdLifeCycleStart = rfiDetail["ProdLifeCycleStart"].ToString();
                            var ProdLifeCycleEnd = rfiDetail["ProdLifeCycleEnd"].ToString();

                            //var RfiDetailStatus = rfiDetail["RfiDetailStatus"].ToString();
                            //, ISNULL(c.UserNo, '') AS SignSalesNo, ISNULL(c.UserName, '') AS SignSalesName, c.Gender AS SignSalesGender
                            //, ISNULL(g.UserNo, '') AS SignSuperSalesNo, ISNULL(g.UserName, '') AS SignSuperSalesName, g.Gender AS SignSuperSalesGender
                            //, ISNULL(h.UserNo, '') AS SignManagerPlanNo, ISNULL(h.UserName, '') AS SignManagerPlanName, h.Gender AS SignManagerPlanGender
                            //, ISNULL(FORMAT(a.ConfirmSSTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmSSTime
                            //, ISNULL(FORMAT(a.ConfirmMPTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmMPTime
                            //, ISNULL(FORMAT(a.ConfirmCustTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmCustTime
                            //, ISNULL(FORMAT(a.ConfirmSTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmSTime
                            //, ISNULL(FORMAT(a.ConfirmFinalTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmFinalTime

                            //string ConfirmSSTime = rfiDetail["ConfirmSSTime"].ToString();
                            //string ConfirmMPTime = rfiDetail["ConfirmMPTime"].ToString();
                            //string ConfirmSTime = rfiDetail["ConfirmSTime"].ToString();
                            string ConfirmFinalTime = rfiDetail["ConfirmFinalTime"].ToString();

                            //string SignSalesName = rfiDetail["SignSalesName"].ToString();
                            //string SignSuperSalesName = rfiDetail["SignSuperSalesName"].ToString();
                            //string SignManagerPlanName = rfiDetail["SignManagerPlanName"].ToString();

                            //string SalesDesc = rfiDetail["SalesDesc"].ToString();
                            //string SuperSalesDesc = rfiDetail["SuperSalesDesc"].ToString();
                            //string ManagerPlanDesc = rfiDetail["ManagerPlanDesc"].ToString();

                            htmlText = htmlText.Replace("[RfiSequence]", RfiSequence);
                            htmlText = htmlText.Replace("[Edition]", Edition);
                            htmlText = htmlText.Replace("[TargetUnitPrice]", TargetUnitPrice.ToString("N"));
                            htmlText = htmlText.Replace("[LifeCycleQty]", LifeCycleQty.ToString("N"));
                            htmlText = htmlText.Replace("[Revenue]", Revenue.ToString("N"));
                            htmlText = htmlText.Replace("[RevenueFC]", RevenueFC.ToString("N"));
                            htmlText = htmlText.Replace("[ExchangeRate]", ExchangeRate.ToString("N"));
                            htmlText = htmlText.Replace("[CurrencyName]", CurrencyName);
                            htmlText = htmlText.Replace("[Currency]", Currency == "NT$" ? "NTD" : Currency);
                            htmlText = htmlText.Replace("[ProdDate]", ProdDate);
                            htmlText = htmlText.Replace("[ProdLifeCycleStart]", ProdLifeCycleStart);
                            htmlText = htmlText.Replace("[ProdLifeCycleEnd]", ProdLifeCycleEnd);

                            //int.TryParse(RfiDetailStatus, out int rfiDetailStatus);

                            if (!rfiDetail["ProductSpec"].ToString().TryParseJson(out JObject resultProdSpec)) throw new SystemException("【產品規格】資料錯誤!");

                            int prodSpec_cnt = resultProdSpec["data"].Count();

                            if (prodSpec_cnt > 0)
                            {
                                int mod = prodSpec_cnt % pageNum;
                                int page = prodSpec_cnt / pageNum + (mod != 0 ? 1 : 0);

                                string htmlDetail = "";
                                string htmlSpec = "<tr>";
                                string htmlFeature = "<tr>";

                                List<string> controlTypes = new List<string> { "R", "C", "D" };

                                int i = 0;
                                foreach (var prodSpec in resultProdSpec["data"].Where(w => controlTypes.Contains(w["ControlType"].ToString())))
                                {
                                    if (prodSpec["ProductSpecDetail"] != null)
                                    {
                                        i++;
                                        htmlSpec += string.Format(@"<td class='data-title'>{0}</td>", prodSpec["SpecName"].ToString().Replace("<", "&lt;").Replace(">", "&gt;"));

                                        if (!prodSpec["ProductSpecDetail"].ToString().TryParseJson(out JObject resultFeature)) throw new SystemException("【產品規格】資料錯誤!");

                                        int feature_cnt = resultFeature["data"].Count();

                                        if (feature_cnt > 0)
                                            htmlFeature += string.Format(@"<td class='text-center'>{0}</td>", 
                                                string.Join("<br />", 
                                                resultFeature["data"]
                                                .Where(w => w["Status"].ToString() == "A")
                                                .Select(s => string.Format("{0}{1}", s["FeatureName"].ToString().Replace("<", "&lt;").Replace(">", "&gt;"), s["Description"]?.ToString() == "" ? "" : ":" + s["Description"]?.ToString()))));
                                        else htmlFeature += "<td class='text-center'>-</td>";

                                        if (i % 7 == 0)
                                        {
                                            htmlDetail += htmlSpec + "</tr>" + htmlFeature + "</tr>";
                                            htmlSpec = "<tr>";
                                            htmlFeature = "<tr>";
                                            i = 0;
                                        }
                                    }
                                }

                                if (i > 0)
                                {
                                    for (int j = 0; j < 7 - i; j++)
                                    {
                                        htmlSpec += "<td class='data-title'></td>";
                                        htmlFeature += "<td class='text-center'>-</td>";
                                    }
                                }

                                htmlDetail += htmlSpec + "</tr>" + htmlFeature + "</tr>";

                                i = 0;
                                controlTypes = new List<string> { "T", "A" };
                                foreach (var prodSpec in resultProdSpec["data"].Where(w => controlTypes.Contains(w["ControlType"].ToString())))
                                {
                                    i++;
                                    htmlSpec = string.Format(@"<td class='data-title'>{0}</td>", prodSpec["SpecName"].ToString());

                                    if (!prodSpec["ProductSpecDetail"].ToString().TryParseJson(out JObject resultFeature)) throw new SystemException("【產品規格】資料錯誤!");

                                    int feature_cnt = resultProdSpec["data"].Count();

                                    if (feature_cnt > 0)
                                    {
                                        string tempData = string.Join("<br />", 
                                            resultFeature["data"]
                                            .Where(w => w["Status"].ToString() == "A")
                                            .Select(s => s["FeatureName"].ToString() + ":" + s["Description"].ToString()));
                                        tempData = regex_newline.Replace(tempData, "<br />");
                                        htmlFeature = string.Format(@"<td colspan='6' class='text-left'>{0}</td>", tempData);
                                    }

                                    htmlDetail += "<tr>" + htmlSpec + htmlFeature + "</tr>";
                                }

                                htmlText = htmlText.Replace("[htmlDetail]", htmlDetail);
                            }

                            if (rfiDetail["SignFlow"].ToString().TryParseJson(out JObject resultSignFlow))
                            {
                                string htmlFlows = string.Empty;
                                int maxCol = 3;
                                int c = 1; //欄位index
                                int r = 0; //列數

                                void GenFlowHeader()
                                {
                                    c = 1; //從第一欄開始產生

                                    var title_row = resultSignFlow["data"].Where(w => w["SortEdition"]?.ToString() == "1").OrderByDescending(o => o["SortNumber"]?.ToString()).Skip(maxCol * r).Take(maxCol).ToList();

                                    if (title_row.Count > 0)
                                    {
                                        string htmlHeader = @"<tr>";
                                        foreach (var title in title_row)
                                        {
                                            if (c == 1)
                                            {
                                                htmlHeader += "<td class='tb2-td' colspan='3'>{0}</td>";
                                            }
                                            else
                                            {
                                                htmlHeader += "<td class='tb2-td' colspan='2'>{0}</td>";
                                            }
                                            htmlHeader = string.Format(htmlHeader, title["SignJobName"]?.ToString());
                                            c++;
                                        }
                                        for (int i = 0; i < maxCol - c + 1; i++)
                                        {
                                            htmlHeader += "<td class='tb2-td' colspan='2'>-</td>";
                                        }
                                        htmlHeader += "</tr>";
                                        htmlFlows += htmlHeader; //add title
                                        r++;
                                    }
                                }

                                int cc = 0;
                                foreach (var flow in resultSignFlow["data"].Where(w => w["SortEdition"]?.ToString() == "1").OrderByDescending(o => o["SortNumber"]?.ToString())) //僅取得有效的流程檔
                                {
                                    string signDesc = flow["SignDesc"]?.ToString();
                                    string signName = flow["FlowerName"]?.ToString();
                                    string signDateTime = flow["SignDateTime"]?.ToString();
                                    string innerHtml = string.Empty;

                                    if (cc % maxCol == 0)
                                    {
                                        cc = 0;
                                        GenFlowHeader(); //產生title
                                        innerHtml += @"<tr>";
                                    }

                                    switch (flow["FlowStatusDesc"]?.ToString())
                                    {
                                        case "Sales":
                                            innerHtml += @"<td class='text-left' colspan='" + (cc == 0 ? "3" : "2") + "'>{1}<br/>建議：{0}<br/>確認時間：{2}</td>";
                                            break;
                                        case "SalesManager":
                                        case "QAManager":
                                        case "Rd":
                                        case "RdManager":
                                        case "Mold&CoatingManager":
                                        case "President":
                                            string MpSignFlag = !string.IsNullOrWhiteSpace(flow["Approved"]?.ToString()) ? flow["ApprovedName"]?.ToString() : string.Empty;
                                            innerHtml += @"<td class='text-left' colspan='" + (cc == 0 ? "3" : "2") + "'>{1}<br/>建議：{0}<br/>簽核時間：{2}</td>";
                                            htmlText = htmlText.Replace("[SignStatus]", MpSignFlag);
                                            htmlText = htmlText.Replace("[SignManagerPlanName]", signName);
                                            break;
                                    }

                                    htmlFlows += string.Format(innerHtml, signDesc, signName, signDateTime);
                                    cc++;
                                }

                                for (int i = 0; i < maxCol - cc; i++)
                                {
                                    htmlFlows += "<td class='text-center' colspan='2'>-</td>";
                                }
                                htmlFlows += "</tr>";

                                htmlText = htmlText.Replace("[htmlSign]", htmlFlows);
                            }

                            logger.Info(string.Format("市場評估單列印(html內容),使用者:{0},RfiId:{1},時間:{2},內容:{3}", requestForInformationDA.CurrentUser, RfiId, DateTime.Now, htmlText));
                        }
                    }
                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream stream = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
                        {
                            PdfWriter pdfWriter = PdfWriter.GetInstance(document, stream); //指定文件預設開檔時的縮放為100%
                            PdfDestination pdfDestination = new PdfDestination(PdfDestination.XYZ, 0, document.PageSize.Height, 1f); //開啟Document文件 

                            document.Open(); //開啟Document文件 

                            XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, document, input, null, Encoding.UTF8, new UnicodeFontFactory()); //使用XMLWorkerHelper把Html parse到PDF檔裡

                            PdfAction action = PdfAction.GotoLocalPage(1, pdfDestination, pdfWriter); //將pdfDest設定的資料寫到PDF檔

                            pdfWriter.SetOpenAction(action);
                        }
                    }

                    pdfFile = stream.ToArray();
                }

                string fileGuid = Guid.NewGuid().ToString();
                Session[fileGuid] = pdfFile;

                #region //條碼圖片刪除
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/RfiBarcode"));
                FileInfo[] files1 = di.GetFiles();
                foreach (FileInfo file1 in files1)
                {
                    file1.Delete();
                }
                #endregion

                #endregion

                #endregion

                logger.Info(string.Format("市場評估單列印(結束),使用者:{0},RfiId:{1},時間:{2}", requestForInformationDA.CurrentUser, RfiId, DateTime.Now));

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = string.Format("RFI市場評估單_{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
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

        #region //GetDesignPdf 取得Design Pdf
        public void GetDesignPdf(int DesignId = -1, string DesignStatus = "", string FlowStatus = "", string OrderBy = "")
        {
            try
            {
                logger.Info(string.Format("設計申請單列印(開始),使用者:{0},DesignId:{1},時間:{2}", requestForInformationDA.CurrentUser, DesignId, DateTime.Now));

                WebApiLoginCheck("DesignManagement", "print");

                Regex regex_newline = new Regex("(\r\n|\r|\n)");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - RFI單頭資料
                requestForInformationDA = new RequestForInformationDA();
                dataRequest = requestForInformationDA.GetDemandDesign(DesignId, -1, "", -1, -1, -1, "", "", "", "", "", DesignStatus, FlowStatus, OrderBy, -1, -1);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                var writer = new BarcodeWriter  //dll裡面可以看到屬性
                {
                    Format = BarcodeFormat.CODE_128,
                    Options = new EncodingOptions //設定大小
                    {
                        Height = 50,
                        Width = 200,
                        PureBarcode = false,
                        Margin = 0
                    }
                };
                
                string htmlText = "";

                if (jsonResponse["status"].ToString() == "success")
                {
                    if (jsonResponse["result"].Count() <= 0) throw new SystemException("【新設計申請單】資料錯誤!");

                    #region //html
                    foreach (var design in jsonResponse["result"])
                    {
                        if (!design["DemandDesignSpec"].ToString().TryParseJson(out JObject resultDesignSpec)) throw new SystemException("【設計規格】資料錯誤!");

                        htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/RfiDemandDesignPdf.html"));

                        #region 產生條碼
                        var img = writer.Write(Convert.ToString(DesignId));
                        Bitmap myBitmap = new Bitmap(img);
                        string FileName = "DesignId" + design["DesignId"].ToString();
                        string filePath = Server.MapPath(string.Format("~/PdfTemplate/MES/RfiBarcode/{0}.bmp", FileName));
                        myBitmap.Save(filePath, ImageFormat.Bmp);
                        #endregion

                        List<(int, string)> DictLogos = new List<(int, string)> { (2, "jmo_logo"), (3, "eterge_logo") };
                        var corp = DictLogos.Find(f => f.Item1.ToString() == design["CompanyId"]?.ToString());
                        if (string.IsNullOrWhiteSpace(corp.Item2)) throw new SystemException("【公司別】資料錯誤!");

                        htmlText = htmlText.Replace("[CompanyLogo]", Server.MapPath(string.Format("~/Areas/SCM/Content/images/{0}.png", corp.Item2)));
                        htmlText = htmlText.Replace("[Barcode]", filePath);

                        var DesignNo = design["DesignNo"].ToString();
                        var DepartmentName = design["DepartmentName"].ToString();
                        var AssemblyName = design["AssemblyName"].ToString();
                        var CreateDate = design["CreateDate"].ToString();
                        //var DesignStatus = rfi["DesignStatus"].ToString();

                        string SalesName = design["SalesName"].ToString();
                        string SuperSalesName = design["SuperSalesName"].ToString();

                        htmlText = htmlText.Replace("[DesignNo]", DesignNo);
                        htmlText = htmlText.Replace("[DepartmentName]", DepartmentName);
                        htmlText = htmlText.Replace("[CreateDate]", CreateDate);
                        htmlText = htmlText.Replace("[SalesName]", SalesName);
                        htmlText = htmlText.Replace("[SuperSalesName]", SuperSalesName);
                        htmlText = htmlText.Replace("[AssemblyName]", AssemblyName);

                        if (design["SignFlow"].ToString().TryParseJson(out JObject signFlows))
                        {
                            string htmlFlows = string.Empty;

                            foreach (var flow in signFlows["data"])
                            {
                                string signDesc = flow["SignDesc"]?.ToString();
                                string signName = flow["FlowerName"]?.ToString();
                                string signDateTime = flow["SignDateTime"]?.ToString();

                                string styleHtml_1 = flow["SortNumber"]?.ToString() == "1" ? "border-width: 3px 3px 0 3px;" : "border-width: 0 3px 0 3px;";
                                string styleHtml_2 = flow["SortNumber"]?.ToString() == "1" ? "border-width: 0 0 3px 3px;" : "border-width: 0 0 3px 3px;";
                                string styleHtml_3 = flow["SortNumber"]?.ToString() == "1" ? "border-width: 0 3px 3px 0;" : "border-width: 0 3px 3px 0;";

                                string styleHtml_4 = flow["SortNumber"]?.ToString() == "1" ? "border-width: 3px 0 0 3px;" : "border-width: 0 0 0 3px;";
                                string styleHtml_5 = flow["SortNumber"]?.ToString() == "1" ? "border-width: 3px 3px 0 0;" : "border-width: 0 3px 0 0;";
                                string innerHtml = string.Empty;

                                switch (flow["FlowStatusDesc"]?.ToString())
                                {
                                    case "Sales":
                                    case "SalesManager":
                                    case "QAManager":
                                        innerHtml = @"<tr style='height:67px;'><td colspan='2' style='vertical-align:top; border: solid; {0}'>建議：{3}</td></tr>
                                                    <tr style='height:29px;'>
                                                        <td class='text-right' style='vertical-align:bottom; border: solid; {1}'></td>
                                                        <td class='text-right' style='vertical-align:bottom; border: solid; {2}'>單位主管：{4} {5}</td>
                                                    </tr>";
                                        htmlFlows += string.Format(innerHtml, styleHtml_1, styleHtml_2, styleHtml_3, signDesc, signName, signDateTime);
                                        break;
                                    case "Rd":
                                    case "RdManager":
                                    case "Mold&CoatingManager":
                                        string SuperRdSignFlag = flow["FlowStatusDesc"]?.ToString() == "RdManager" ? flow["Approved"]?.ToString() == "N" ? "不可行" : (flow["Approved"]?.ToString() == "Y" ? "可行" : string.Empty) : string.Empty;
                                        innerHtml = @"<tr style='height:67px;'><td colspan='2' style='vertical-align:top; border: solid; {0}'>建議：{3}</td></tr>
                                                    <tr style='height:29px;'>
                                                        <td class='text-right' style='vertical-align:bottom; border: solid; {1}'>評估結果：{6}</td>
                                                        <td class='text-right' style='vertical-align:bottom; border: solid; {2}'>單位主管：{4} {5}</td>
                                                    </tr>";
                                        htmlFlows += string.Format(innerHtml, styleHtml_1, styleHtml_2, styleHtml_3, signDesc, signName, signDateTime, SuperRdSignFlag);
                                        break;
                                    case "President":
                                        string MpSignFlag = !string.IsNullOrWhiteSpace(flow["Approved"]?.ToString()) ? flow["ApprovedName"]?.ToString() : string.Empty;
                                        innerHtml = @"<tr style='height:67px;'>
                                                        <td style='vertical-align:top; border: solid; {0}'>建議：{4}</td>
                                                        <td class='text-right' style='vertical-align:top; border: solid; {1}'><h3>{7}</h3></td>
                                                    </tr>
                                                    <tr style='height:29px;'>
                                                        <td style='vertical-align:top; border: solid; {2}'><h3>總經理同意後，始可投入資源進行開發。</h3></td>
                                                        <td class='text-right' style='vertical-align:bottom; border: solid; {3}'><h3>總經理：{5}</h3>{6}</td>
                                                    </tr>";
                                        htmlFlows += string.Format(innerHtml, styleHtml_4, styleHtml_5, styleHtml_2, styleHtml_3, signDesc, signName, signDateTime, MpSignFlag);
                                        break;
                                }
                            }
                            htmlText = htmlText.Replace("[htmlFlows]", htmlFlows);
                        }

                        //int.TryParse(DesignStatus, out int designStatus);
                        //var SuperRdSignFlag = designStatus == 7 && !string.IsNullOrWhiteSpace(SignSuperRdName) && string.IsNullOrWhiteSpace(ConfirmMPTime) ? "不可行" : "可行";
                        //var MpSignFlag = designStatus == 7 && !string.IsNullOrWhiteSpace(SignManagerPlanName) && string.IsNullOrWhiteSpace(ConfirmFinalTime) ? "不准" : (designStatus == 6 ? "核准" : "");

                        //string SignSuperSalesName = rfi["SignSuperSalesName"].ToString();
                        //string SignSuperRdName = rfi["SignSuperRdName"].ToString();
                        //string SignManagerPlanName = rfi["SignManagerPlanName"].ToString();

                        //string ConfirmSRDTime = rfi["ConfirmSRDTime"].ToString();
                        //string ConfirmMPTime = rfi["ConfirmMPTime"].ToString();
                        //string ConfirmFinalTime = rfi["ConfirmFinalTime"].ToString();

                        //string SuperSalesDesc = rfi["SuperSalesDesc"].ToString();
                        //string SuperRdDesc = rfi["SuperRdDesc"].ToString();
                        //string ManagerPlanDesc = rfi["ManagerPlanDesc"].ToString();

                        //htmlText = htmlText.Replace("[SuperSalesDesc]", SuperSalesDesc);
                        //htmlText = htmlText.Replace("[SuperSalesName]", SignSuperSalesName);
                        //htmlText = htmlText.Replace("[ConfirmSRDTime]", ConfirmSRDTime);

                        //htmlText = htmlText.Replace("[SuperRdDesc]", SuperRdDesc);
                        //htmlText = htmlText.Replace("[SuperRdName]", SignSuperRdName);
                        //htmlText = htmlText.Replace("[ConfirmMPTime]", ConfirmMPTime);

                        //htmlText = htmlText.Replace("[ManagerPlanDesc]", ManagerPlanDesc);
                        //htmlText = htmlText.Replace("[ManagerPlanName]", SignManagerPlanName);
                        //htmlText = htmlText.Replace("[ConfirmFinalTime]", ConfirmFinalTime);

                        //int pageNum = 15;

                        string htmlTemplate = htmlText;

                        string htmlDetail = "";
                        string htmlSpec = "";
                        string htmlFeature = "";

                        List<string> controlTypes = new List<string> { "R", "C", "D", "T", "A" };
                        List<string> controlTypeCRD = new List<string> { "R", "C", "D" };

                        int i = 0;
                        foreach (var designSpec in resultDesignSpec["data"].Where(w => controlTypes.Contains(w["ControlType"].ToString())))
                        {
                            i++;
               
                            htmlSpec = string.Format("<tr><td colspan='6' style='font-size:14px; font-weight: 600;'>{0}. {1}</td></tr>", i, designSpec["ParameterName"].ToString());

                            if (designSpec["DemandDesignSpecDetail"] != null)
                            {
                                htmlFeature = "";
                                if (!designSpec["DemandDesignSpecDetail"].ToString().TryParseJson(out JObject resultPmtDetail)) throw new SystemException("【設計規格特徵】資料錯誤!");
                                foreach (var designSpecDetail in resultPmtDetail["data"])
                                {
                                    string designInput = designSpecDetail["DesignInput"].ToString().Replace("<", "&lt;").Replace(">", "&gt;");
                                    string description = designSpecDetail["Description"].ToString().Replace("<", "&lt;").Replace(">", "&gt;");

                                    htmlFeature += string.Format("<tr><td>{0}</td><td class='tb4-c1'>{1}</td><td class='tb4-c1'>{2}</td><td class='tb4-c1 text-right'>{3}</td></tr>",
                                        controlTypeCRD.Contains(designSpec["ControlType"].ToString()) ? string.Empty : designSpecDetail["PmtDetailName"].ToString(),
                                        designSpecDetail["RequireFlag"].ToString() == "0" ? "v" : "",
                                        controlTypeCRD.Contains(designSpec["ControlType"].ToString()) ? designSpecDetail["PmtDetailName"].ToString() : designInput, 
                                        description);
                                }

                                htmlDetail += htmlSpec + htmlFeature;
                            }
                        }

                        htmlText = htmlText.Replace("[htmlDetail]", htmlDetail);

                        logger.Info(string.Format("設計申請單列印(html內容),使用者:{0},DesignId:{1},時間:{2},內容:{3}", requestForInformationDA.CurrentUser, DesignId, DateTime.Now, htmlText));
                    }
                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream stream = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
                        {
                            PdfWriter pdfWriter = PdfWriter.GetInstance(document, stream); //指定文件預設開檔時的縮放為100%
                            PdfDestination pdfDestination = new PdfDestination(PdfDestination.XYZ, 0, document.PageSize.Height, 1f); //開啟Document文件 

                            document.Open(); //開啟Document文件 

                            XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, document, input, null, Encoding.UTF8, new UnicodeFontFactory()); //使用XMLWorkerHelper把Html parse到PDF檔裡

                            PdfAction action = PdfAction.GotoLocalPage(1, pdfDestination, pdfWriter); //將pdfDest設定的資料寫到PDF檔

                            pdfWriter.SetOpenAction(action);
                        }
                    }

                    pdfFile = stream.ToArray();
                }

                string fileGuid = Guid.NewGuid().ToString();
                Session[fileGuid] = pdfFile;

                #region //條碼圖片刪除
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/RfiBarcode"));
                FileInfo[] files1 = di.GetFiles();
                foreach (FileInfo file1 in files1)
                {
                    file1.Delete();
                }
                #endregion

                #endregion

                #endregion

                logger.Info(string.Format("設計申請單列印(結束),使用者:{0},DesignId:{1},時間:{2}", requestForInformationDA.CurrentUser, DesignId, DateTime.Now));

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = string.Format("RFI新設計需求申請單_{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
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