using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog.Internal;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Ocsp;
using SCMDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class PurchaseOrderController : WebController
    {
        private PurchaseOrderDA purchaseOrderDA = new PurchaseOrderDA();

        #region //View
        public ActionResult MaterialManagement()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult PurchaseOrderManagement()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetMaterial 取得材料資料
        [HttpPost]
        public void GetMaterial(int MtId = -1, string MaterialName = "", string MtlItemNo = "", int SupplierId = -1, string Status = "", string CanUse =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetMaterial(MtId, MaterialName, MtlItemNo, SupplierId, Status, CanUse
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

        #region //GetMtAmount 取得材料金額資料
        [HttpPost]
        public void GetMtAmount(int MtId = -1, int MaId = -1, string ConfirmStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MaterialManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetMtAmount(MtId, MaId, ConfirmStatus
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

        #region //GetMtAmountERP 取得採購材料金額(ERP)
        [HttpPost]
        public void GetMtAmountERP(string MaterialName = "")
        {
            try
            {
                WebApiLoginCheck("PurchaseOrderManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetMtAmountERP(MaterialName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPurchaseOrder 取得採購單資料
        [HttpPost]
        public void GetPurchaseOrder(int PoId = -1, string PoErpFullNo = "", int SupplierId = -1, int ConfirmUserId = -1, string SearchKey = ""
            , string ConfirmStatus = "", string ClosureStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseOrderManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPurchaseOrder(PoId, PoErpFullNo, SupplierId, ConfirmUserId, SearchKey
                    , ConfirmStatus, ClosureStatus, StartDate, EndDate
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

        #region //GetPoDetail 取得採購單身資料
        [HttpPost]
        public void GetPoDetail(int PoDetailId = -1, int PoId = -1, string PoErpFullNo = "", int SupplierId = -1, int ConfirmUserId = -1, string SearchKey = ""
            , string ConfirmStatus = "", string ClosureStatus = "", string StartDate = "", string EndDate = "", string PaymentTermNo = "", string CurrencyCode = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseOrderManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPoDetail(PoDetailId, PoId, PoErpFullNo, SupplierId, ConfirmUserId, SearchKey
                    , ConfirmStatus, ClosureStatus, StartDate, EndDate, PaymentTermNo, CurrencyCode
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

        #region//Add
        #region//AddMaterial 新增材料資料
        [HttpPost]
        public void AddMaterial(string MaterialName = "", int SupplierId = -1, string Remark = "", int MtlItemId = -1)
        {
            try
            {
                
                WebApiLoginCheck("MaterialManagement", "add");


                #region //Request
                dataRequest = purchaseOrderDA.AddMaterial(MaterialName, SupplierId, Remark, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddMtAmount 新增材料金額資料
        [HttpPost]
        public void AddMtAmount(int MtId = -1, float Amount = -1, float UnitQuantity = -1)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "add");


                #region //Request
                dataRequest = purchaseOrderDA.AddMtAmount(MtId, Amount, UnitQuantity);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddPurchaseOrder 採購單自動開立
        [HttpPost]
        [Route("api/ERP/PurchaseOrder")]
        public void AddPurchaseOrder(string CompanyNo = "", string SecretKey = "",string UserNo="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "AddPurchaseOrder");
                #endregion

                #region //Request
                dataRequest = purchaseOrderDA.AddPurchaseOrder(CompanyNo, UserNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region//UpdateMaterial 更新材料資料
        [HttpPost]
        public void UpdateMaterial(int MtId, string MaterialName = "", int SupplierId = -1, string Remark = "", int MtlItemId = -1)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "update");


                #region //Request
                dataRequest = purchaseOrderDA.UpdateMaterial(MtId, MaterialName, SupplierId, Remark, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateMaterialStatus 更新材料狀態資料
        [HttpPost]
        public void UpdateMaterialStatus(int MtId)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "status-switch");


                #region //Request
                dataRequest = purchaseOrderDA.UpdateMaterialStatus(MtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateMtAmount 更新材料金額資料
        [HttpPost]
        public void UpdateMtAmount(int MaId, double Amount = -1, double UnitQuantity = -1)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "update");


                #region //Request
                dataRequest = purchaseOrderDA.UpdateMtAmount(MaId, Amount, UnitQuantity);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateMtAmountConfirm 更新確認材料金額資料
        [HttpPost]
        public void UpdateMtAmountConfirm(int MaId)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "confirm");


                #region //Request
                dataRequest = purchaseOrderDA.UpdateMtAmountConfirm(MaId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region//DeleteMaterial 刪除材料資料
        [HttpPost]
        public void DeleteMaterial(int MtId)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "delete");


                #region //Request
                dataRequest = purchaseOrderDA.DeleteMaterial(MtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteMtAmount 刪除材料金額資料
        [HttpPost]
        public void DeleteMtAmount(int MaId)
        {
            try
            {

                WebApiLoginCheck("MaterialManagement", "delete");


                #region //Request
                dataRequest = purchaseOrderDA.DeleteMtAmount(MaId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //API FPR BM
        #region //UpdatePurchaseOrderSynchronize 採購單資料同步
        [HttpPost]
        [Route("api/ERP/PurchaseOrderSynchronize")]
        public void UpdatePurchaseOrderSynchronize(string Company, string SecretKey, string UpdateDate, string PoErpFullNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdatePurchaseOrderSynchronize");
                #endregion

                #region //Request
                dataRequest = purchaseOrderDA.UpdatePurchaseOrderSynchronize(Company, UpdateDate, PoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //API FOR SRM
        #region //GetCompanyKanbanInfo 取得公司看板資料
        [HttpPost]
        [Route("api/SCM/GetCompanyKanbanInfo")]
        public void GetCompanyKanbanInfo(string CompanyNo = "", string NewCompanyNo = "", string SupplierNos = "", string SecretKey = "")
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetCompanyKanbanInfo");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetCompanyKanbanInfo(CompanyNo, NewCompanyNo, SupplierNos);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTaxTypeERP 取得ERP稅別資料
        [HttpPost]
        [Route("api/SCM/GetTaxTypeERP")]
        public void GetTaxTypeERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string TaxCode = "", string TaxName = "", string TypeNo = "", string Taxation = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetTaxType");

                if (CompanyNo.Length > 0) 
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetTaxTypeERP(SupplierCompany
                    , TaxCode, TaxName, TypeNo, Taxation, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTradeTermERP 取得ERP交易條件
        [HttpPost]
        [Route("api/SCM/GetTradeTermERP")]
        public void GetTradeTermERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetTradeTerm");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetTradeTermERP(SupplierCompany, IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPaymentTermERP 取得ERP付款條件資料
        [HttpPost]
        [Route("api/SCM/GetPaymentTermERP")]
        public void GetPaymentTermERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string PaymentType = "", string PaymentTermCode = "", string PaymentTermCodes = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPaymentTerm");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPaymentTermERP(SupplierCompany
                    , PaymentType, PaymentTermCode, PaymentTermCodes, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCurrencyInfo 取得ERP幣別資料
        [HttpPost]
        [Route("api/SCM/GetCurrencyInfo")]
        public void GetCurrencyInfo(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string Currency = "", string CurrencyName = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetCurrency");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetCurrencyInfo(SupplierCompany
                    , Currency, CurrencyName, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetExchangeRate 取得ERP匯率資料
        [HttpPost]
        [Route("api/SCM/GetExchangeRate")]
        public void GetExchangeRate(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string Currency = "", string StartDate = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1, bool TopOne = false
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetExchangeRate");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetExchangeRate(SupplierCompany
                    , Currency, StartDate, SearchKey
                    , IsOptions, draw, TopOne
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetItemUnitConvertERP 取得ERP幣別資料
        [HttpPost]
        [Route("api/SCM/GetItemUnitConvertERP")]
        public void GetItemUnitConvertERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string MtlItemNo = "", string MtlItemNos = "", string ConvertUnit = "", string Unit = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetItemUnitConvert");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetItemUnitConvertERP(SupplierCompany
                    , MtlItemNo, MtlItemNos, ConvertUnit, Unit, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSupplierERP 取得ERP供應商資料
        [HttpPost]
        [Route("api/SCM/GetSupplierERP")]
        public void GetSupplierERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string SupplierNo = "", string SupplierNos = ""
            , string LastTradingStartDate = "", string LastTradingEndDate = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetSupplier");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo, SupplierNo = SupplierNo, SupplierNos = SupplierNos });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetSupplierERP(SupplierCompany
                    , LastTradingStartDate, LastTradingEndDate, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCompanyERP 取得ERP公司資料
        [HttpPost]
        [Route("api/SCM/GetCompanyERP")]
        public void GetCompanyERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string CompanyName = "", string CompanyEnglishName = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetCustomer");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetCompanyERP(SupplierCompany
                    , SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlItemERP 取得ERP品號資料
        [HttpPost]
        [Route("api/SCM/GetMtlItemERP")]
        public void GetMtlItemERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string MtlItemNo = "", string MtlItemNos = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetMtlItem");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetMtlItemERP(SupplierCompany
                    , MtlItemNo, MtlItemNos, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPurchaseRequisitionERP 取得ERP請購單頭資料
        [HttpPost]
        [Route("api/SCM/GetPurchaseRequisitionERP")]
        public void GetPurchaseRequisitionERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string PrErpPrefix = "", string PrErpNo = "", string PrErpFullNo = ""
            , string SupplierNo = "", string SupplierNos = ""
            , string PrErpPrefixs = "", string PrErpNos = ""
            , string PrUser = "", string MtlItemNo = "", string MtlItemNos = ""
            , string PrStartDate = "", string PrEndDate = ""
            , string DocStartDate = "", string DocEndDate = ""
            , string ConfirmStatus = "", string ClosureStatus = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPurchaseRequisition");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPurchaseRequisitionERP(SupplierCompany
                    , PrErpPrefix, PrErpNo, PrErpFullNo
                    , SupplierNo, SupplierNos
                    , PrErpPrefixs, PrErpNos
                    , PrUser, MtlItemNo, MtlItemNos
                    , PrStartDate, PrEndDate
                    , DocStartDate, DocEndDate
                    , ConfirmStatus, ClosureStatus, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrDetailERP 取得ERP請購單身資料
        [HttpPost]
        [Route("api/SCM/GetPrDetailERP")]
        public void GetPrDetailERP(List<SupplierCompany> SupplierCompany
            , string CompanyNo = "", string NewCompanyNo = "", string SupplierNo = "", string SupplierNos = "", string MtlItemNo = "", string InventoryNos = "", string Currency = ""
            , string PrErpPrefix = "", string PrErpNo = "", string PrSeq = ""
            , string PrErpPrefixs = "", string PrErpNos = "", string PrSeqs = ""
            , string PrUser = "", string PoUsers = "", string DepartmentNo = "", string LockStatus = ""
            , string PrStartDate = "", string PrEndDate = ""
            , string DemandStartDate = "", string DemandEndDate = ""
            , string DeliveryStartDate = "", string DeliveryEndDate = ""
            , string ConfirmStatus = "", string ClosureStatus = "", bool? DocTransfer = null, string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPrDetail");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPrDetailERP(SupplierCompany
                    , SupplierNo, SupplierNos, MtlItemNo, InventoryNos, Currency
                    , PrErpPrefix, PrErpNo, PrSeq
                    , PrErpPrefixs, PrErpNos, PrSeqs
                    , PrUser, PoUsers, DepartmentNo, LockStatus
                    , PrStartDate, PrEndDate
                    , DemandStartDate, DemandEndDate
                    , DeliveryStartDate, DeliveryEndDate
                    , ConfirmStatus, ClosureStatus, DocTransfer, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPurchaseOrderERP 取得ERP採購單頭資料
        [HttpPost]
        [Route("api/SCM/GetPurchaseOrderERP")]
        public void GetPurchaseOrderERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string SupplierNo = "", string SupplierNos = ""
            , string PoErpPrefix = "", string PoErpNo = "", string PoErpFullNo = ""
            , string PoErpPrefixs = "", string PoErpNos = "", string PoErpFullNos = ""
            , string PoUser = "", string MtlItemNo = "", string MtlItemNos = "", string Currency = ""
            , string PoStartDate = "", string PoEndDate = ""
            , string DocStartDate = "", string DocEndDate = ""
            , string ConfirmStatus = "", string ClosureStatus = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPurchaseOrder");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null) 
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPurchaseOrderERP(SupplierCompany
                    , SupplierNo, SupplierNos
                    , PoErpPrefix, PoErpNo, PoErpFullNo
                    , PoErpPrefixs, PoErpNos, PoErpFullNos
                    , PoUser, MtlItemNo, MtlItemNos, Currency
                    , PoStartDate, PoEndDate
                    , DocStartDate, DocEndDate
                    , ConfirmStatus, ClosureStatus, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPoDetailERP 取得ERP採購單身資料
        [HttpPost]
        [Route("api/SCM/GetPoDetailERP")]
        public void GetPoDetailERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string SupplierNo = "", string SupplierNos = ""
            , string PoErpPrefix = "", string PoErpNo = "", string PoSeq = ""
            , string PoErpPrefixs = "", string PoErpNos = "", string PoSeqs = ""
            , string PoErpFullNo = "", string PoErpFullNos = ""
            , string PoErpPrefixNo = "", string PoErpPrefixNos = ""
            , string PromiseStartDate = "", string PromiseEndDate = ""
            , string OriPromiseStartDate = "", string OriPromiseEndDate = ""
            , string DeliveryStartDate = "", string DeliveryEndDate = ""
            , string ConfirmStatus = "", string ClosureStatus = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPoDetail");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPoDetailERP(SupplierCompany
                    , SupplierNo, SupplierNos
                    , PoErpPrefix, PoErpNo, PoSeq
                    , PoErpPrefixs, PoErpNos, PoSeqs
                    , PoErpFullNo, PoErpFullNos
                    , PoErpPrefixNo, PoErpPrefixNos
                    , PromiseStartDate, PromiseEndDate
                    , OriPromiseStartDate, OriPromiseEndDate
                    , DeliveryStartDate, DeliveryEndDate
                    , ConfirmStatus, ClosureStatus, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPriceCheckingERP 取得ERP核價單頭資料
        [HttpPost]
        [Route("api/SCM/GetPriceCheckingERP")]
        public void GetPriceCheckingERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string PcErpPrefix = "", string PcErpNo = ""
            , string PcErpPrefixs = "", string PcErpNos = ""
            , string Currency = "", string TaxIncluded = ""
            , string PcStartDate = "", string PcEndDate = ""
            , string DocStartDate = "", string DocEndDate = ""
            , string ConfirmStatus = "", string DocStatus = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPriceChecking");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPriceCheckingERP(SupplierCompany
                    , PcErpPrefix, PcErpNo
                    , PcErpPrefixs, PcErpNos
                    , Currency, TaxIncluded
                    , PcStartDate, PcEndDate
                    , DocStartDate, DocEndDate
                    , ConfirmStatus, DocStatus, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPcDetailERP 取得ERP核價單身資料
        [HttpPost]
        [Route("api/SCM/GetPcDetailERP")]
        public void GetPcDetailERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string PcErpPrefix = "", string PcErpNo = "", string PcSeq = ""
            , string PcErpPrefixs = "", string PcErpNos = "", string PcSeqs = ""
            , string MtlItemNo = "", string MtlItemNos = "", string MtlItemName = ""
            , string SupplierItemNo = "", string SupplierItemNos = "", string UnitPricing = ""
            , string EffectiveStartDate = "", string EffectiveEndDate = ""
            , string ExpirationStartDate = "", string ExpirationEndDate = ""
            , string ConfirmStatus = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetPcDetail");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPcDetailERP(SupplierCompany
                    , PcErpPrefix, PcErpNo, PcSeqs
                    , PcErpPrefixs, PcErpNos, PcSeqs
                    , MtlItemNo, MtlItemNos, MtlItemName
                    , SupplierItemNo, SupplierItemNos, UnitPricing
                    , EffectiveStartDate, EffectiveEndDate
                    , ExpirationStartDate, ExpirationEndDate
                    , ConfirmStatus, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetGrDetailERP 取得ERP進貨單身
        [HttpPost]
        [Route("api/SCM/GetGrDetailERP")]
        public void GetGrDetailERP(List<SupplierCompany> SupplierCompany
            , string CompanyNo = "", string NewCompanyNo = "", string GrErpPrefix = "", string GrErpNo = "", string GrSeq = "", string GrErpFullNo = "", string GrErpPrefixNo = ""
            , string GrErpFullNos = "", string GrErpPrefixNos = ""
            , string PoErpPrefix = "", string PoErpNo = "", string PoSeq = "", string PoErpFullNo = "", string PoErpPrefixNo = ""
            , string PoErpFullNos = "", string PoErpPrefixNos = ""
            , string MtlItemNo = "", string MtlItemNos = "", string MtlItemName = ""
            , string AcceptanceStartDate = "", string AcceptanceEndDate = ""
            , string EffectiveStartDate = "", string EffectiveEndDate = ""
            , string ConfirmStatus = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetGrDetail");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetGrDetailERP(SupplierCompany
                    , GrErpPrefix, GrErpNo, GrSeq, GrErpFullNo, GrErpPrefixNo
                    , GrErpFullNos, GrErpPrefixNos
                    , PoErpPrefix, PoErpNo, PoSeq, PoErpFullNo, PoErpPrefixNo
                    , PoErpFullNos, PoErpPrefixNos
                    , MtlItemNo, MtlItemNos, MtlItemName
                    , AcceptanceStartDate, AcceptanceEndDate
                    , EffectiveStartDate, EffectiveEndDate
                    , ConfirmStatus, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRgDetailERP 取得ERP退貨單身資料
        [HttpPost]
        [Route("api/SCM/GetRgDetailERP")]
        public void GetRgDetailERP(List<SupplierCompany> SupplierCompany = null
            , string CompanyNo = "", string NewCompanyNo = "", string RgErpPrefix = "", string RgErpNo = "", string RgSeq = ""
            , string RgErpFullNo = "", string RgErpPrefixNo = ""
            , string RgErpFullNos = "", string RgErpPrefixNos = ""
            , string PoErpPrefix = "", string PoErpNo = "", string PoSeq = ""
            , string PoErpFullNo = "", string PoErpPrefixNo = ""
            , string PoErpFullNos = "", string PoErpPrefixNos = ""
            , string MtlItemNo = "", string MtlItemName = "", string SupplierNo = ""
            , string ConfirmStatus = "", string CheckOutCode = "", string SearchKey = ""
            , bool IsOptions = false, string SecretKey = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "GetRgDetail");

                if (CompanyNo.Length > 0)
                {
                    if (SupplierCompany == null)
                    {
                        SupplierCompany = new List<SupplierCompany>();
                        SupplierCompany.Add(new SupplierCompany { CompanyNo = CompanyNo, NewCompanyNo = NewCompanyNo });
                    }
                    else SupplierCompany = SupplierCompany.Where(w => w.CompanyNo == CompanyNo).ToList();
                }

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetRgDetailERP(SupplierCompany
                    , RgErpPrefix, RgErpNo, RgSeq
                    , RgErpFullNo, RgErpPrefixNo
                    , RgErpFullNos, RgErpPrefixNos
                    , PoErpPrefix, PoErpNo, PoSeq
                    , PoErpFullNo, PoErpPrefixNo
                    , PoErpFullNos, PoErpPrefixNos
                    , MtlItemNo, MtlItemName, SupplierNo
                    , ConfirmStatus, CheckOutCode, SearchKey
                    , IsOptions, draw
                    , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePurchaseOrderBatch 更新ERP採購單批次資料
        [HttpPost]
        [Route("api/SCM/UpdatePurchaseOrderBatch")]
        public void UpdatePurchaseOrderBatch(PurchaseOrderMaster model
            , string SecretKey)
        {
            try
            {
                ApiKeyVerify("JMO", SecretKey, "UpdatePurchaseOrder");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.UpdatePurchaseOrderBatch(model);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //Shintokuro
        #region //Get
        #region //GetPurchaseRequisitionBM 取得請購單(BM)(Api)
        [HttpPost]
        [Route("api/SCM/GetPurchaseRequisitionBM")]
        public void GetPurchaseRequisitionBM(string CompanyNo = "", string PrErpPrefix = "", string PrErpNo = ""
            , string PrId = "", string PurchaseOrde = "", string MtlItemNo = "", string MtlItemName = "", string UserNo = "", string SupplierNo = ""
            , string OrderBy = "", string PageIndex = "", string PageSize = "", string SearchKey = ""
            , string SecretKey = "", int draw = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseRequisition");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPurchaseRequisitionBM(CompanyNo, PrErpPrefix, PrErpNo
                    , PrId, PurchaseOrde, MtlItemNo, MtlItemName, UserNo, SupplierNo
                    , SearchKey, OrderBy, PageIndex, PageSize
                    , draw);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrDepartmentERP 取得請購部門(Api)
        [HttpPost]
        [Route("api/SCM/GetPrDepartmentERP")]
        public void GetPrDepartmentERP(string CompanyNo = "", string DepartmentNo = "", string PrUser = "", string PoUser = "", string SearchKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPrDepartmentERP(CompanyNo, DepartmentNo, PrUser, PoUser, SearchKey);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrUserERP 取得請購人員(Api)
        [HttpPost]
        [Route("api/SCM/GetPrUserERP")]
        public void GetPrUserERP( string CompanyNo = "", string UserNo = "",string LockStatus = "", string SearchKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPrUserERP(CompanyNo, UserNo, LockStatus, SearchKey);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPoUserERP 取得採購人員(Api)
        [HttpPost]
        [Route("api/SCM/GetPoUserERP")]
        public void GetPoUserERP(string CompanyNo = "", string UserNo = "", string LockStatus = "", string SearchKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPoUserERP(CompanyNo, UserNo, LockStatus, SearchKey);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPrSupplierErp 取得請購供應商(Api)
        [HttpPost]
        [Route("api/SCM/GetPrSupplierErp")]
        public void GetPrSupplierErp(string SupplierNo = "", string SupplierName = "", string CompanyNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPrSupplierErp(SupplierNo, SupplierName, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPoSupplierERP 取得請購人員
        [HttpPost]
        [Route("api/SCM/GetPoSupplierERP")]
        public void GetPoSupplierERP(string SupplierNo = "", string SupplierName = "", string CompanyNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPoSupplierERP(SupplierNo, SupplierName, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUnitERP 取得單位
        [HttpPost]
        [Route("api/SCM/GetUnitERP")]
        public void GetUnitERP(string UomNo = "", string UomName = "", string UomType = "", string CompanyNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetUnitERP(UomNo, UomName, UomType, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetInventoryERP 取得庫別
        [HttpPost]
        [Route("api/SCM/GetInventoryERP")]
        public void GetInventoryERP(string InventoryNo = "", string InventoryName = "", string CompanyNo = "", string InventoryType = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetInventoryERP(InventoryNo, InventoryName, CompanyNo, InventoryType);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCurrencyErp 取得幣別(Api)
        [HttpPost]
        [Route("api/SCM/GetCurrencyErp")]
        public void GetCurrencyErp(string CompanyNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetCurrencyErp(CompanyNo);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPoErpPrefix 取得稅別碼(Api)
        [HttpPost]
        [Route("api/SCM/GetPoErpPrefix")]
        public void GetPoErpPrefix(string CompanyNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPoErpPrefix(CompanyNo);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion

        #region Add
        #region//AddPoByPrBatch 請購轉採購 (Api)
        [HttpPost]
        [Route("api/SCM/AddPoByPrBatch")]
        public void AddPoByPrBatch(string CompanyNo = ""
            , string PoErpPrefix = "", string PoDate = "", string SupplierNo = ""
            , List<PrDetail> PrItems = null, string CurrentUser = ""
            , string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify("JMO", SecretKey, "AddPurchaseOrder");
                #endregion

                #region //Request
                dataRequest = purchaseOrderDA.AddPoByPrBatch(CompanyNo
                    , PoErpPrefix, PoDate, SupplierNo
                    , PrItems, CurrentUser);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region Update
        #region //UpdatePrDetailERP 更新ERP請購單身(Api)
        [HttpPost]
        [Route("api/SCM/UpdatePrDetailERP")]
        public void UpdatePrDetailERP(string PrErpPrefix = "", string PrErpNo = "", string PrSeq = ""
            , string UrgentMtl = "", string SupplierNo = "", string TradeTerm = "", string TaxCode = "", string Taxation = "", string PoCurrency = "", string PoCurrencyLocal = ""
            , string PoUser = "", string Inventory = "", int PoQty = 0, string PoUnit = ""
            , int PoPriceQty = 0, string PoPriceUnit = "", double PoUnitPrice = 0d, double PoPrice = 0d, string PoRemark = ""
            , string apiUrl = "", string SecretKey = "", string CompanyNo = "", string CurrentUser = ""
            )
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.UpdatePrDetailERP(PrErpPrefix, PrErpNo, PrSeq
                    , UrgentMtl, SupplierNo, TradeTerm, TaxCode, Taxation, PoCurrency, PoCurrencyLocal
                    , PoUser, Inventory, PoQty, PoUnit
                    , PoPriceQty, PoPriceUnit, PoUnitPrice, PoPrice, PoRemark
                    , CompanyNo, CurrentUser);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrDetailCloseERP 更新ERP請購單指定結案(Api)
        [HttpPost]
        [Route("api/SCM/UpdatePrDetailCloseERP")]
        public void UpdatePrDetailCloseERP(string PrErpPrefix = "", string PrErpNo = "", string PrSeq = ""
            , string apiUrl = "", string SecretKey = "", string CompanyNo = "", string CurrentUser = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.UpdatePrDetailCloseERP(PrErpPrefix, PrErpNo, PrSeq
                   , CompanyNo, CurrentUser);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePrDetailReCloseERP 更新ERP請購單反確認指定結案(Api)
        [HttpPost]
        [Route("api/SCM/UpdatePrDetailReCloseERP")]
        public void UpdatePrDetailReCloseERP(string PrErpPrefix = "", string PrErpNo = "", string PrSeq = ""
            , string apiUrl = "", string SecretKey = "", string CompanyNo = "", string CurrentUser = ""
            )
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.UpdatePrDetailReCloseERP(PrErpPrefix, PrErpNo, PrSeq
                   , CompanyNo, CurrentUser);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePoConfirmERP 更新ERP請購單確認(Api)
        [HttpPost]
        [Route("api/SCM/UpdatePoConfirmERP")]
        public void UpdatePoConfirmERP(string PoErpPrefix = "", string PoErpNo = ""
            , string apiUrl = "", string SecretKey = "", string CompanyNo = "", string CurrentUser = ""
            )
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.UpdatePoConfirmERP(PoErpPrefix, PoErpNo
                   , CompanyNo, CurrentUser
                    );
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePoReConfirmERP 更新ERP採購單反確認(Api)
        [HttpPost]
        [Route("api/SCM/UpdatePoReConfirmERP")]
        public void UpdatePoReConfirmERP(string PoErpPrefix = "", string PoErpNo = ""
            , string apiUrl = "", string SecretKey = "", string CompanyNo = "", string CurrentUser = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(CompanyNo, SecretKey, "GetPurchaseOrder");
                #endregion

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.UpdatePoReConfirmERP(PoErpPrefix, PoErpNo, CompanyNo, CurrentUser);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    error = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion
        #endregion
        #endregion
    }
}