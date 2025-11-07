using Helpers;
using SCMDA;
using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System.Dynamic;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;
using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Converters;
using System.Collections.Generic;
using MESDA;

namespace Business_Manager.Controllers
{
    public class SaleOrderController : WebController
    {
        private SaleOrderDA saleOrderDA = new SaleOrderDA();
        private PurchaseOrderDA purchaseOrderDA = new PurchaseOrderDA();

        #region //View
        public ActionResult SaleOrderManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PMDOrderManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetSaleOrder 取得訂單單頭資料
        [HttpPost]
        public void GetSaleOrder(int SoId = -1, string SoErpPrefix = "", string SoErpNo = "", string SoErpFullNo = "", int CustomerId = -1
            , int SalesmenId = -1, string MtlItemNo = "", string StartDate = "", string EndDate = "", string ConfirmStatus = "", string ClosureStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrder(SoId, SoErpPrefix, SoErpNo, SoErpFullNo, CustomerId
                    , SalesmenId, MtlItemNo, StartDate, EndDate, ConfirmStatus, ClosureStatus
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

        #region //GetSaleOrder02 取得訂單資料
        [HttpPost]
        public void GetSaleOrder02(int SoId = -1, string SoErpPrefix = "", string SoErpNo = "", int CustomerId = -1
            , int SalesmenId = -1, string SearchKey = "", string ConfirmStatus = "", string ClosureStatus = ""
            , string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrder02(SoId, SoErpPrefix, SoErpNo, CustomerId
                    , SalesmenId, SearchKey, ConfirmStatus, ClosureStatus
                    , StartDate, EndDate
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

        #region //GetSoDetail 取得訂單單身資料
        [HttpPost]
        public void GetSoDetail(int SoDetailId = -1, int SoId = -1, int MtlItemId = -1, string SoErpFullNo = "", string CustomerMtlItemNo = "", string TransferStatus = "", string SearchKey = "", string MtlItemIdIsNull = "", int CompanyId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoDetail(SoDetailId, SoId, MtlItemId, SoErpFullNo, CustomerMtlItemNo, TransferStatus, SearchKey, MtlItemIdIsNull, CompanyId
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

        #region //GetSoDetail02 取得訂單單身資料
        [HttpPost]
        public void GetSoDetail02(int SoDetailId = -1, int SoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoDetail02(SoDetailId, SoId
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

        #region //GetSaleOrderSimple 取得訂單單頭資料
        [HttpPost]
        public void GetSaleOrderSimple(int MtlItemId)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrderSimple(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //GetTotal 取得訂單單身加總資料
        [HttpPost]
        public void GetTotal(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetTotal(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //GetSoSequence 取得訂單單身流水號資料
        [HttpPost]
        public void GetSoSequence(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoSequence(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion  

        #region //GetSoModification 取得訂單變更單頭資料
        [HttpPost]
        public void GetSoModification(int SomId = -1, int SoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "modify");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoModification(SomId, SoId
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

        #region //GetSoErpFullNo 取得訂單單別+單號
        [HttpPost]
        public void GetSoErpFullNo(string DeliveryStatus = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoErpFullNo(DeliveryStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSoDeposit 取得訂單訂金資料
        [HttpPost]
        public void GetSoDeposit(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoDeposit(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSoDetailTxAmount 試算訂單單身金額與稅額
        [HttpPost]
        public void GetSoDetailTxAmount(int SoId = -1, double UnitPrice = -1, double SoPriceQty = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoDetailTxAmount(SoId, UnitPrice, SoPriceQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSoDetailTemp 取得採購單匯入資料
        [HttpPost]
        public void GetSoDetailTemp(int SoDetailTempId = -1, int SoId = -1, int SoDetailId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoDetailTemp(SoDetailTempId, SoId, SoDetailId, Status
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

        #region //GetPurchaseOrderFromERP 取得ERP採購單身資料
        [HttpPost]
        public void GetPurchaseOrderFromERP(int CompanyId = -1, int CustomerId = -1, string PoErpPrefix = "", string PoErpNo = "", string SearchKey = ""
            , string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseOrderManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPurchaseOrderFromERP(CompanyId, CustomerId, PoErpPrefix, PoErpNo, SearchKey
                    , StartDate, EndDate
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

        #region //GetPurchaseOrderTypeFromERP 取得ERP採購單別資料
        [HttpPost]
        public void GetPurchaseOrderTypeFromERP(int CompanyId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("PurchaseOrderManagement", "read,constrained-data");

                #region //Request
                purchaseOrderDA = new PurchaseOrderDA();
                dataRequest = purchaseOrderDA.GetPurchaseOrderTypeFromERP(CompanyId, SearchKey
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

        #region //GetQuotationErp 取得ERP採購單別資料
        [HttpPost]
        public void GetQuotationErp(int SoId = -1, int MtlItemId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetQuotationErp(SoId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPMDOrderData -- 取得 PMD 訂單資料
        [HttpPost]
        public void GetPMDOrderData(int TypeId = -1 , string Company = "", int CustomerId = -1)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetPMDOrderData(TypeId, CustomerId, Company);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPMDOrderId 取得 PMD 訂單ID
        [HttpPost]
        public void GetPMDOrderId(int TypeId)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetPMDOrderId(TypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPMDOrderId 取得 PMD 訂單ID
        [HttpPost]
        public void GetCurrentProcess(int MtlItemId)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetCurrentProcess(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCurrentEdit
        [HttpPost]
        public void GetCurrentEdit(string PageNames)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetCurrentEdit(PageNames);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCustomer
        [HttpPost]
        public void GetCustomer(int CustomerId = -1, string CustomerNo = "", string CustomerName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetCustomer(CustomerId, CustomerNo, CustomerName
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

        #region //GetPMDPage
        [HttpPost]
        public void GetPMDPage()
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetPMDPage();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPMDOrderCustomerList 取得PMD客戶列表
        [HttpPost]
        public void GetPMDOrderCustomerList(int TypeId = -1)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read,constrained-data");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetPMDOrderCustomerList(TypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddSaleOrder 訂單單頭資料新增
        [HttpPost]
        public void AddSaleOrder(int SoId = -1, int DepartmentId = -1, string SoErpPrefix = "", string SoErpNo = "", string SoDate = ""
            , string DocDate = "", string SoRemark = "", int CustomerId = -1, int SalesmenId = -1, string CustomerAddressFirst = "", string CustomerAddressSecond = ""
            , string CustomerPurchaseOrder = "", string DepositPartial = "", double DepositRate = 0.0, string Currency = "", double ExchangeRate = 0.0
            , string TaxNo = "", string Taxation = "", double BusinessTaxRate = 0.0, string DetailMultiTax = "", double TotalQty = 0.0, double Amount = 0.0
            , double TaxAmount = 0.0, string ShipMethod = "", string TradeTerm = "", string PaymentTerm = "", string PriceTerm = "", string ConfirmStatus = ""
            , int ConfirmUserId = -1, string TransferStatus = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "add");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.AddSaleOrder(SoId, DepartmentId, SoErpPrefix, SoErpNo, SoDate
                    , DocDate, SoRemark, CustomerId, SalesmenId, CustomerAddressFirst, CustomerAddressSecond
                    , CustomerPurchaseOrder, DepositPartial, DepositRate, Currency, ExchangeRate
                    , TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TotalQty, Amount
                    , TaxAmount, ShipMethod, TradeTerm, PaymentTerm, PriceTerm, ConfirmStatus
                    , ConfirmUserId, TransferStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddSaleOrder02 訂單資料新增
        public void AddSaleOrder02(string SoErpPrefix = "", string DocDate = "", int CustomerId = -1
            , string CustomerPurchaseOrder = "", int DepartmentId = -1, int SalesmenId = -1, string SoRemark = "", string Currency = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "add");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.AddSaleOrder02(SoErpPrefix, DocDate, CustomerId
                    , CustomerPurchaseOrder, DepartmentId, SalesmenId, SoRemark, Currency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddSoDetail 訂單單身資料新增
        [HttpPost]
        public void AddSoDetail(int SoId = -1, int SoDetailTempId = -1, string SoSequence = "", int MtlItemId = -1, string SoMtlItemName = ""
            , int InventoryId = -1, double SoQty = -1, double UnitPrice = 0.0, double Amount = 0.0
            , string ProductType = "", float FreebieOrSpareQty = 0, string PromiseDate = ""
            , string Project = "", string SoDetailRemark = "", string QuotationErp = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "add");

                bool updateError = false;
                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.AddSoDetail(SoId, SoDetailTempId, SoSequence, MtlItemId, SoMtlItemName
                    , InventoryId, SoQty, UnitPrice, Amount
                    , ProductType, FreebieOrSpareQty, PromiseDate
                    , Project, SoDetailRemark, QuotationErp);

                jsonResponse = JObject.Parse(dataRequest);
                updateError = jsonResponse["status"].ToString() != "success";
                #endregion

                #region //判斷是否拋轉
                if (!updateError)
                {
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetSoStatus(SoId);

                    jsonResponse = JObject.Parse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest);
                        if (result["data"].Count() <= 0) throw new SystemException("訂單狀態錯誤!");
                        if (result["data"][0]["TransferStatus"].ToString() == "Y")
                        {
                            #region //訂單拋轉ERP
                            saleOrderDA = new SaleOrderDA();
                            dataRequest = saleOrderDA.UpdateSaleOrderImport(SoId);
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("目前尚無資料!");
                    }
                }
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddSoModification 訂單變更單頭資料新增
        [HttpPost]
        public void AddSoModification(int SomId = -1, int SoId = -1, string Version = "", int DepartmentId = -1, string DocDate = "", string SoRemark = ""
            , int SalesmenId = -1, string CustomerAddressFirst = "", string CustomerAddressSecond = "", string CustomerPurchaseOrder = ""
            , string DepositPartial = "", double DepositRate = 0.0, string Currency = "", double ExchangeRate = 0.0, string TaxNo = "", string Taxation = ""
            , double BusinessTaxRate = 0.0, string DetailMultiTax = "", string ShipMethod = "", string TradeTerm = "", string PaymentTerm = "", string PriceTerm = ""
            , string ClosureStatus = "", string ModiReason = "", string ConfirmStatus = "", int ConfirmUserId = -1, string TransferStatus = "", string TransferDate = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "modify");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.AddSoModification(SomId, SoId, Version, DepartmentId, DocDate, SoRemark
                    , SalesmenId, CustomerAddressFirst, CustomerAddressSecond, CustomerPurchaseOrder
                    , DepositPartial, DepositRate, Currency, ExchangeRate, TaxNo, Taxation
                    , BusinessTaxRate, DetailMultiTax, ShipMethod, TradeTerm, PaymentTerm, PriceTerm
                    , ClosureStatus, ModiReason, ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPMDOrderLog -- 新增PMD訂單LOG紀錄
        [HttpPost]
        public void AddPMDOrderLog(int UserId = -1, int EditItem = -1, string EditTable = "")

        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "modify");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.AddPMDOrderLog(UserId, EditItem, EditTable);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateSaleOrder 訂單單頭資料更新
        [HttpPost]
        public void UpdateSaleOrder(int SoId = -1, int DepartmentId = -1, string SoErpPrefix = "", string SoErpNo = "", string SoDate = ""
            , string DocDate = "", string SoRemark = "", int CustomerId = -1, int SalesmenId = -1, string CustomerAddressFirst = "", string CustomerAddressSecond = ""
            , string CustomerPurchaseOrder = "", string DepositPartial = "", double DepositRate = 0.0, string Currency = "", double ExchangeRate = 0.0
            , string TaxNo = "", string Taxation = "", double BusinessTaxRate = 0.0, string DetailMultiTax = "", double TotalQty = 0.0, double Amount = 0.0
            , double TaxAmount = 0.0, string ShipMethod = "", string TradeTerm = "", string PaymentTerm = "", string PriceTerm = "", string ConfirmStatus = ""
            , int ConfirmUserId = -1, string TransferStatus = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrder(SoId, DepartmentId, SoErpPrefix, SoErpNo, SoDate
                    , DocDate, SoRemark, CustomerId, SalesmenId, CustomerAddressFirst, CustomerAddressSecond
                    , CustomerPurchaseOrder, DepositPartial, DepositRate, Currency, ExchangeRate
                    , TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TotalQty, Amount
                    , TaxAmount, ShipMethod, TradeTerm, PaymentTerm, PriceTerm, ConfirmStatus
                    , ConfirmUserId, TransferStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrder02 訂單資料更新
        public void UpdateSaleOrder02(int SoId = -1, string SoErpPrefix = "", string DocDate = "", int CustomerId = -1
            , string CustomerPurchaseOrder = "", int DepartmentId = -1, int SalesmenId = -1, string SoRemark = "", string Currency = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "update");

                bool updateError = false;
                #region //資料更新
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrder02(SoId, SoErpPrefix, DocDate, CustomerId
                    , CustomerPurchaseOrder, DepartmentId, SalesmenId, SoRemark, Currency);

                jsonResponse = JObject.Parse(dataRequest);
                updateError = jsonResponse["status"].ToString() != "success";
                #endregion

                #region //判斷是否拋轉
                if (!updateError)
                {
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetSoStatus(SoId);

                    jsonResponse = JObject.Parse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest);
                        if (result["data"].Count() <= 0) throw new SystemException("訂單狀態錯誤!");
                        if (result["data"][0]["TransferStatus"].ToString() == "Y")
                        {
                            #region //訂單拋轉ERP
                            saleOrderDA = new SaleOrderDA();
                            dataRequest = saleOrderDA.UpdateSaleOrderImport(SoId);
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("目前尚無資料!");
                    }
                }
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSoDetail 訂單單身資料更新
        [HttpPost]
        public void UpdateSoDetail(int SoId = -1, int SoDetailId = -1, int SoDetailTempId = -1, string SoSequence = "", int MtlItemId = -1, string SoMtlItemName = ""
            , int InventoryId = -1, double SoQty = -1, double UnitPrice = 0.0, double Amount = 0.0
            , string ProductType = "", float FreebieOrSpareQty = 0, string PromiseDate = ""
            , string Project = "", string SoDetailRemark = "", string QuotationErp = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "update");

                bool updateError = false;
                #region //資料更新
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSoDetail(SoDetailId, SoDetailTempId, SoSequence, MtlItemId, SoMtlItemName
                    , InventoryId, SoQty, UnitPrice, Amount
                    , ProductType, FreebieOrSpareQty, PromiseDate
                    , Project, SoDetailRemark, QuotationErp);

                jsonResponse = JObject.Parse(dataRequest);
                updateError = jsonResponse["status"].ToString() != "success";
                #endregion

                #region //判斷是否拋轉
                if (!updateError)
                {
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetSoStatus(SoId);

                    jsonResponse = JObject.Parse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest);
                        if (result["data"].Count() <= 0) throw new SystemException("訂單狀態錯誤!");
                        if (result["data"][0]["TransferStatus"].ToString() == "Y")
                        {
                            #region //訂單拋轉ERP
                            saleOrderDA = new SaleOrderDA();
                            dataRequest = saleOrderDA.UpdateSaleOrderImport(SoId);
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("目前尚無資料!");
                    }
                }
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderImport 訂單拋轉ERP
        [HttpPost]
        public void UpdateSaleOrderImport(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "import");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderImport(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderConfirm 訂單ERP確認
        [HttpPost]
        public void UpdateSaleOrderConfirm(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "confirm");

                bool importError = false;
                #region //判斷是否拋轉
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoStatus(SoId);

                jsonResponse = JObject.Parse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("訂單狀態錯誤!");
                    if (result["data"][0]["TransferStatus"].ToString() == "N")
                    {
                        #region //訂單拋轉ERP
                        saleOrderDA = new SaleOrderDA();
                        dataRequest = saleOrderDA.UpdateSaleOrderImport(SoId);

                        jsonResponse = JObject.Parse(dataRequest);
                        importError = jsonResponse["status"].ToString() != "success";
                        #endregion
                    }
                }
                else
                {
                    throw new SystemException("目前尚無資料!");
                }
                #endregion

                if (!importError)
                {
                    #region //確認
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.UpdateSaleOrderConfirm(SoId);
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

        #region //UpdateSaleOrderReConfirm 訂單ERP反確認
        [HttpPost]
        public void UpdateSaleOrderReConfirm(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "reconfirm");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderReConfirm(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderVoid 訂單ERP作廢
        [HttpPost]
        public void UpdateSaleOrderVoid(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "void");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderVoid(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderModify 訂單資料變更
        [HttpPost]
        public void UpdateSaleOrderModify(string SoDetails = "", string Deposits = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "modify");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderModify(SoDetails, Deposits);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderToERP 訂單資料拋轉
        [HttpPost]
        public void UpdateSaleOrderToERP(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "confirm");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderToERP(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSoModification 訂單變更單頭資料更新
        [HttpPost]
        public void UpdateSoModification(int SomId = -1, int SoId = -1, string Version = "", int DepartmentId = -1, string DocDate = "", string SoRemark = ""
            , int SalesmenId = -1, string CustomerAddressFirst = "", string CustomerAddressSecond = "", string CustomerPurchaseOrder = ""
            , string DepositPartial = "", double DepositRate = 0.0, string Currency = "", double ExchangeRate = 0.0, string TaxNo = "", string Taxation = ""
            , double BusinessTaxRate = 0.0, string DetailMultiTax = "", string ShipMethod = "", string TradeTerm = "", string PaymentTerm = "", string PriceTerm = ""
            , string ClosureStatus = "", string ModiReason = "", string ConfirmStatus = "", int ConfirmUserId = -1, string TransferStatus = "", string TransferDate = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "modify");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSoModification(SomId, SoId, Version, DepartmentId, DocDate, SoRemark
                    , SalesmenId, CustomerAddressFirst, CustomerAddressSecond, CustomerPurchaseOrder
                    , DepositPartial, DepositRate, Currency, ExchangeRate, TaxNo, Taxation
                    , BusinessTaxRate, DetailMultiTax, ShipMethod, TradeTerm, PaymentTerm, PriceTerm
                    , ClosureStatus, ModiReason, ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderStatus 訂單資料狀態更新
        [HttpPost]
        public void UpdateSaleOrderStatus(int SoId = -1, string StatusName = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "reconfirm");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderStatus(SoId, StatusName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSaleOrderManualSynchronize 訂單資料手動同步
        [HttpPost]
        public void UpdateSaleOrderManualSynchronize(string SoErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "sync");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSaleOrderManualSynchronize(SoErpFullNo, SyncStartDate, SyncEndDate
                    , NormalSync, TranSync);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSoTransferBpm 拋轉訂單單據至BPM(若ERP未拋，連同ERP一起拋)
        [HttpPost]
        public void UpdateSoTransferBpm(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "bpm-transfer");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSoTransferBpm(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSoDetailTempBatch --取得採購單匯入資料 -- Chia Yuan 2024.04.18
        [HttpPost]
        public void UpdateSoDetailTempBatch(int SoId = -1, int CompanyId = -1, string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateSoDetailTempBatch(SoId, CompanyId, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UploadPmdOrderData -- 上傳PMD資料
        [HttpPost]
        public void UploadPmdOrderData(string PmdJsonData = "", int TypeId = -1, int customerId = -1)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UploadPmdOrderData(PmdJsonData, TypeId, customerId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePageLock  -- 鎖定頁更新 
        [HttpPost]
        public void UpdatePageLock(string PageNames, int isLock)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdatePageLock(PageNames, isLock);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePageUnLock  -- 鎖定頁解鎖 
        [HttpPost]
        public void UpdatePageUnLock(string PageNames)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdatePageUnLock(PageNames);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePageSended  -- PMD修改為已送
        public void UpdatePageSended(string SendedIds, string pmdOrdertempIds)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "update");   

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdatePageSended(SendedIds, pmdOrdertempIds);
                #endregion

                #region //Response
                var result = BaseHelper.DAResponse(dataRequest);
                logger.Info($"UpdatePageSended Result: {result}");
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                logger.Error("UpdatePageSended failed: " + e.Message);
                #endregion

            }

            //Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateOrderToPMDOrder 訂單更新至PMD進度表
        public void UpdateOrderToPMDOrder(int TypeId)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UpdateOrderToPMDOrder(TypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UploadPmdOrderDataTemp -- 上傳PMD暫存資料
        [HttpPost]
        public void UploadPmdOrderDataTemp(string PmdJsonData = "", int TypeId = -1, int customerId = -1)
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "update");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.UploadPmdOrderDataTemp(PmdJsonData, TypeId, customerId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteSaleOrder 訂單資料刪除
        [HttpPost]
        public void DeleteSaleOrder(int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "delete");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.DeleteSaleOrder(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteSoDetail 訂單單身資料刪除
        [HttpPost]
        public void DeleteSoDetail(int SoDetailId = -1, int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "delete");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.DeleteSoDetail(SoDetailId, SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteSoDetail02 訂單單身資料刪除
        [HttpPost]
        public void DeleteSoDetail02(int SoDetailId = -1, int SoId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "delete");

                bool updateError = false;
                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.DeleteSoDetail02(SoDetailId);

                jsonResponse = JObject.Parse(dataRequest);
                updateError = jsonResponse["status"].ToString() != "success";
                #endregion

                #region //判斷是否拋轉
                if (!updateError)
                {
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetSoStatus(SoId);

                    jsonResponse = JObject.Parse(dataRequest);
                    if (jsonResponse["status"].ToString() == "success")
                    {
                        var result = JObject.Parse(dataRequest);
                        if (result["data"].Count() <= 0) throw new SystemException("訂單狀態錯誤!");
                        if (result["data"][0]["TransferStatus"].ToString() == "Y")
                        {
                            #region //訂單拋轉ERP
                            saleOrderDA = new SaleOrderDA();
                            dataRequest = saleOrderDA.UpdateSaleOrderImport(SoId);
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("目前尚無資料!");
                    }
                }
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteSoModification 訂單變更單頭資料刪除
        [HttpPost]
        public void DeleteSoModification(int SomId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "modify");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.DeleteSoModification(SomId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteSoDetailTemp 採購單匯入資料刪除
        [HttpPost]
        public void DeleteSoDetailTemp(int SoId = -1, int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "delete");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.DeleteSoDetailTemp(SoId, CompanyId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateSaleOrderSynchronize 訂單資料同步
        [HttpPost]
        [Route("api/ERP/SaleOrderSynchronize")]
        public void UpdateSaleOrderSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateSaleOrderSynchronize");
                #endregion

                #region //Request
                dataRequest = saleOrderDA.UpdateSaleOrderSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBpmSoStatusData -- 取得BPM訂單回傳狀態資料
        [HttpPost]
        public void GetBpmSoStatusData(string id = "", string status = "", string rootId = "", string comfirmUser = "", string bpmNo = "", string company = "")
        {
            try
            {
                string ErpFlag = "P";
                var dataRequestJson = new JObject();

                #region //先記錄LOG
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.AddSoBpmLog(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser);
                dataRequestJson = JObject.Parse(dataRequest);
                int SoBpmLogId = -1;
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    foreach (var item in dataRequestJson["data"])
                    {
                        SoBpmLogId = Convert.ToInt32(item["SoBpmLogId"]);
                    }
                }
                #endregion

                //若狀態為Y，則將ERP訂單進行核單動作
                if (status == "Y")
                {
                    #region //核准ERP訂單
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.UpdateSaleOrderConfirmByBpm(Convert.ToInt32(id), company, comfirmUser);
                    #endregion

                    dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["status"].ToString() != "success")
                    {
                        ErpFlag = "F";

                        #region //將錯誤訊息回寫LOG
                        dataRequest = saleOrderDA.UpdateSoBpmLogErrorMessage(SoBpmLogId, dataRequestJson["msg"].ToString());
                        JObject logDataRequestJson = JObject.Parse(dataRequest);
                        if (logDataRequestJson["status"].ToString() != "success")
                        {
                            throw new SystemException(logDataRequestJson["msg"].ToString());
                        }
                        #endregion
                    }
                }

                #region //更改MES訂單狀態
                dataRequest = saleOrderDA.UpdateSoBpmInfo(Convert.ToInt32(id), bpmNo, status, rootId, comfirmUser, ErpFlag);
                #endregion

                if (ErpFlag == "F")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
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

        #region //CheckCreditLimit 檢核客戶信用額度
        [HttpPost]
        [Route("api/BM/CheckCreditLimit")]
        public void CheckCreditLimit(string CustomerNo = "", string Currency = "", decimal TotalAmount = -1, string DocType = "", decimal Amount = -1, string CompanyNo = "")
        {
            try
            {
                //WebApiLoginCheck("SaleOrderManagement", "read");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.CheckCreditLimit(CustomerNo, Currency, TotalAmount, DocType, Amount, CompanyNo);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(dataRequest);
        }
        #endregion

        #region //GetSaleOrderEIP -- 取得客戶訂單資料
        [HttpPost]
        [Route("api/SCM/GetSaleOrderEIP")]
        public void GetSaleOrderEIP(int SoId = -1, string SoErpNo = "", string SoErpFullNo = "", string CustomerPurchaseOrder = "", string SearchKey = "", string StartDate = "", string EndDate = ""
            , int MemberId = -1, int[] CustomerIds = null
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrderEIP(SoId, SoErpNo, SoErpFullNo, CustomerPurchaseOrder, SearchKey, StartDate, EndDate
                    , MemberId, CustomerIds
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

        #region //GetSoDetailEIP --取得客戶訂單身資料
        [HttpPost]
        [Route("api/SCM/GetSoDetailEIP")]
        public void GetSoDetailEIP(int SoDetailId = -1, int SoId = -1, string SoErpFullNo = "", string TransferStatus = "", string SearchKey = ""
            , int MemberId = -1, int[] CustomerIds = null
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSoDetailEIP(SoDetailId, SoId, SoErpFullNo, TransferStatus, SearchKey
                    , MemberId, CustomerIds
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

        #region //GetTotalEIP --取得訂單單身加總資料
        [HttpPost]
        [Route("api/SCM/GetTotalEIP")]
        public void GetTotalEIP(int SoId = -1)
        {
            try
            {
                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetTotal(SoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetPendingSaleOrderEIP --未結訂單
        [HttpPost]
        [Route("api/SCM/GetPendingSaleOrder")]
        public void GetPendingSaleOrder(string Company = "", string SecretKey = ""
            , string OrderDate = "", string CustomerNo = "", string CustomerShortName = "", string CustomerOrderNo = "", string Currency = "", string ExchangeRate = "", string SalesmenName = ""
            , string MtlItemNo = "", string MtlItemName = "", string PromiseDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "PendingSaleOrder");
                #endregion

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetPendingSaleOrder(Company, OrderDate, CustomerNo, CustomerShortName, CustomerOrderNo, Currency, ExchangeRate, SalesmenName
                    , MtlItemNo, MtlItemName, PromiseDate
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

        #region //SendPMDOrderExcels 寄送PMD
        [HttpPost]
        public void SendPMDOrderExcels(string SendCustomer, string SendExcel)
        {
            switch (SendExcel)
            {
                case "ViewPMDOrder1":
                    SendPMDOrderExcel1(SendCustomer);
                    break;
                case "ViewPMDOrder2":
                    SendPMDOrderExcel2(SendCustomer);
                    break;
                case "ViewPMDOrder3":
                    SendPMDOrderExcel3(SendCustomer);
                    break;
                case "ViewPMDOrder4":
                    SendPMDOrderExcel4(SendCustomer);
                    break;
                case "ViewPMDOrder5":
                    SendPMDOrderExcel5(SendCustomer);
                    break;
            }



        }
        #endregion

        #region PMDAlertMamo 進度表推送MAMO異常通知 
        public void PMDAlertMamo(string Company, string PageName, string ItemName, string CustomerMtlItem, string MtlItemName)
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "OspMAMO");
                #endregion

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.PMDAlertMamo(Company, PageName, ItemName, CustomerMtlItem, MtlItemName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

           // Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//SendPMDOrderExcel 寄送PMD-1
        [HttpPost]
        [Route("api/SCM/SendPMDOrderExcel1")]
        public void SendPMDOrderExcel1(string SendCustomer = "", string Company = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "SendPMDOrderExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion
                string[] sendCustomerList = SendCustomer.Split(',');
                foreach (var customer in sendCustomerList)
                {
                    int customerId = int.Parse(customer);

                    #region //Request
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetPMDOrderDataSend(1, customerId, Company);

                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                    {

                        byte[] excelBytes;
                        string fileName = $"模芯衬装配件-{DateTime.Now:yyyyMMdd}.xlsx";

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
                        defaultStyle.Font.FontName = "微軟正黑體";
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
                        titleStyle.Font.FontName = "微軟正黑體";
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
                        headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                        headerStyle.Font.FontName = "微軟正黑體";
                        headerStyle.Font.FontSize = 14;
                        headerStyle.Font.Bold = true;
                        #endregion

                        #region //dataStyle
                        var dataStyle = XLWorkbook.DefaultStyle;
                        dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.TopBorderColor = XLColor.Black;
                        dataStyle.Border.BottomBorderColor = XLColor.Black;
                        dataStyle.Border.LeftBorderColor = XLColor.Black;
                        dataStyle.Border.RightBorderColor = XLColor.Black;
                        dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                        dataStyle.Font.FontName = "微軟正黑體";
                        dataStyle.Font.FontSize = 12;
                        dataStyle.Font.Bold = false;
                        #endregion

                        #region //numberStyle
                        var numberStyle = XLWorkbook.DefaultStyle;
                        numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
                        numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                        numberStyle.Font.FontName = "微軟正黑體";
                        numberStyle.Font.FontSize = 12;
                        numberStyle.Font.Bold = false;
                        numberStyle.NumberFormat.Format = "#,##0.00";
                        #endregion

                        #region //currencyStyle
                        var currencyStyle = XLWorkbook.DefaultStyle;
                        currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "微軟正黑體";
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.NumberFormat.Format = "#,##0.00";
                        #endregion
                        #endregion

                        #region //EXCEL

                        using (var workbook = new XLWorkbook())
                        {

                            var worksheet = workbook.Worksheets.Add("模芯衬装配件");
                            worksheet.Style = defaultStyle;

                            // 設定標題
                            worksheet.Cell("A1").Value = "模芯衬装配件";
                            var titleRange = worksheet.Range("A1:T1");
                            titleRange.Merge().Style = titleStyle;
                            worksheet.Row(1).Height = 40;

                            // 設定表頭
                            string[] headers = new string[] { "廠內品號", "模號/品號", "品名", "訂單數量", "備品數量", "接單日期", "下圖日", "當前進度", "客戶交期", "回覆交期", "延誤交期時間"
                            , "備註", "發貨地", "半成品發貨(ZY->JC)", "預留欄位1", "預留欄位2", "預留欄位3", "預留欄位4" };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = worksheet.Cell(2, i + 1);
                                cell.Value = headers[i];
                                cell.Style = headerStyle;
                            }

                            // 寫入數據
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                            int rowIndex = 2;
                            string pmdOrderIds = string.Join(",", data.Select(d => d["PMDOrderId"].ToString()));
                            string pmdOrdertempIds = string.Join(",", data.Select(d => d["PMDOrderTempId"].ToString()));


                            #region //BODY
                            if (data.Count() > 0)
                            {

                                foreach (var item in data)
                                {
                                    if (item.CustomerId == null || item.CustomerId <= 0) throw new Exception("客戶欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.MtlItemNo == null || item.MtlItemNo == "") throw new Exception("品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.CustomerMtlItemNo == null || item.CustomerMtlItemNo == "") throw new Exception("模號/品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.MtlItemName == null || item.MtlItemNo == "") throw new Exception("品名欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.SoQty == null || item.SoQty == 0) throw new Exception("訂單數量不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.OrderDate == null || item.OrderDate == "") throw new Exception("接單日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.CustomerDueDate == null || item.CustomerDueDate == "") throw new Exception("客戶交期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ConfirmedDueDate == null || item.ConfirmedDueDate == "") throw new Exception("回覆交期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ShipFrom == null || item.ShipFrom == "") throw new Exception("發貨地欄位不得為空! 資料序號: " + item.PMDOrderTempId);


                                    rowIndex++;

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CustomerMtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemName.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.SoQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.SpareQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.OrderDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.DrawingAttachmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.CurrentProcess.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.CustomerDueDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ConfirmedDueDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.DelayRemark.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.Remark.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.ShipFrom.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.SemiDelivery.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.PMDItem01.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.PMDItem02.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.PMDItem03.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).Value = item.PMDItem04.ToString().Trim();

                                    //rowIndex++;


                                }

                                #endregion

                                #region //設定

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion

                                #region 輸出檔案
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    workbook.SaveAs(ms);
                                    excelBytes = ms.ToArray();
                                }
                                #endregion

                                #region//Mail 寄送托外進貨週期性報表

                                //取得Mail樣板設定
                                string SettingSchema = "SendPMDOrderExcel1";
                                string SettingNo = "Y";

                                var saleOrderDA = new SaleOrderDA();
                                string mailSettingsJson = saleOrderDA.SendCustomerMail(customerId);
                                if (string.IsNullOrEmpty(mailSettingsJson))
                                {
                                    throw new Exception("無法取得郵件設定");
                                }

                                var parsedSettings = JObject.Parse(mailSettingsJson);
                                if (parsedSettings["status"].ToString() != "success")
                                {
                                    throw new Exception(parsedSettings["msg"].ToString());
                                }

                                dynamic mailSettings = parsedSettings["data"];

                                foreach (var item in mailSettings)
                                {
                                    var mailFile = new MailFile
                                    {
                                        FileName = Path.GetFileNameWithoutExtension(fileName),
                                        FileExtension = ".xlsx",
                                        FileContent = excelBytes
                                    };


                                    #region //寄送Mail
                                    var mailConfig = new MailConfig
                                    {
                                        Host = item.Host,
                                        Port = Convert.ToInt32(item.Port),
                                        SendMode = Convert.ToInt32(item.SendMode),
                                        From = item.MailFrom,
                                        Subject = $"模芯衬装配件_{DateTime.Now:yyyy/MM/dd}",
                                        Account = item.Account,
                                        Password = item.Password,
                                        MailTo = item.SalesMail,
                                        MailCc = item.SalesMailCc,
                                        MailBcc = item.SalesMailBcc,
                                        HtmlBody = $"附件為模芯衬装配件，日期:{DateTime.Now:yyyy/MM/dd}",
                                        TextBody = "-",
                                        FileInfo = new List<MailFile> { mailFile },
                                        QcFileFlag = "N"
                                    };
                                    BaseHelper.MailSend(mailConfig);
                                    #endregion
                                }
                                #endregion

                                UpdatePageSended(pmdOrderIds, pmdOrdertempIds);

                                #region //Response
                                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "寄送完成"
                                });
                                #endregion

                            }
                            else
                            {
                                PMDAlertMamo(Company, "模芯衬装配件", "該客戶無資料", "", "");

                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "error",
                                    msg = "該客戶無資料"
                                });
                                #endregion

                                logger.Error("該客戶無資料");

                            }


                            #endregion


                        }
                        #endregion


                    }
                    else
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = JObject.Parse(dataRequest)["msg"].ToString()
                        });
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                PMDAlertMamo(Company, "模芯衬装配件", e.Message, "", "");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//SendPMDOrderExcel 寄送PMD-2
        [HttpPost]
        [Route("api/SCM/SendPMDOrderExcel2")]
        public void SendPMDOrderExcel2(string SendCustomer = "", string Company = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "SendPMDOrderExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion
                string[] sendCustomerList = SendCustomer.Split(',');
                foreach (var customer in sendCustomerList)
                {
                    int customerId = int.Parse(customer);
                    #region //Request
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetPMDOrderDataSend(2, customerId, Company);

                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                    {

                        byte[] excelBytes;
                        string fileName = $"精定位订单明细-{DateTime.Now:yyyyMMdd}.xlsx";

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
                        defaultStyle.Font.FontName = "微軟正黑體";
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
                        titleStyle.Font.FontName = "微軟正黑體";
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
                        headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                        headerStyle.Font.FontName = "微軟正黑體";
                        headerStyle.Font.FontSize = 14;
                        headerStyle.Font.Bold = true;
                        #endregion

                        #region //dataStyle
                        var dataStyle = XLWorkbook.DefaultStyle;
                        dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.TopBorderColor = XLColor.Black;
                        dataStyle.Border.BottomBorderColor = XLColor.Black;
                        dataStyle.Border.LeftBorderColor = XLColor.Black;
                        dataStyle.Border.RightBorderColor = XLColor.Black;
                        dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                        dataStyle.Font.FontName = "微軟正黑體";
                        dataStyle.Font.FontSize = 12;
                        dataStyle.Font.Bold = false;
                        #endregion

                        #region //numberStyle
                        var numberStyle = XLWorkbook.DefaultStyle;
                        numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
                        numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                        numberStyle.Font.FontName = "微軟正黑體";
                        numberStyle.Font.FontSize = 12;
                        numberStyle.Font.Bold = false;
                        numberStyle.NumberFormat.Format = "#,##0.00";
                        #endregion

                        #region //currencyStyle
                        var currencyStyle = XLWorkbook.DefaultStyle;
                        currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "微軟正黑體";
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.NumberFormat.Format = "#,##0.00";
                        #endregion
                        #endregion

                        #region //EXCEL

                        using (var workbook = new XLWorkbook())
                        {

                            var worksheet = workbook.Worksheets.Add("精定位订单明细");
                            worksheet.Style = defaultStyle;

                            // 設定標題
                            worksheet.Cell("A1").Value = "精定位订单明细";
                            var titleRange = worksheet.Range("A1:O1");
                            titleRange.Merge().Style = titleStyle;
                            worksheet.Row(1).Height = 40;

                            // 設定表頭
                            string[] headers = new string[] { "廠內品號", "模號/品號", "接單日", "訂單數量", "預計出貨日", "預計出貨數量", "實際出貨日", "實際出貨數量", "備註", "發貨地"
                            , "預留欄位1", "預留欄位2", "預留欄位3", "預留欄位4" };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = worksheet.Cell(2, i + 1);
                                cell.Value = headers[i];
                                cell.Style = headerStyle;
                            }

                            // 寫入數據
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                            int rowIndex = 2;
                            string pmdOrderIds = string.Join(",", data.Select(d => d["PMDOrderId"].ToString()));
                            string pmdOrdertempIds = string.Join(",", data.Select(d => d["PMDOrderTempId"].ToString()));

                            #region //BODY
                            if (data.Count() > 0)
                            {
                                foreach (var item in data)
                                {
                                    if (item.CustomerId == null || item.CustomerId == 0) throw new Exception("客戶欄位不得為空! 資料序號: " + item.PMDOrderTempId);

                                    if (item.MtlItemNo == null || item.MtlItemNo == "") throw new Exception("廠內品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.CustomerMtlItemNo == null || item.CustomerMtlItemNo == "") throw new Exception("模號/品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.SoQty == null || item.SoQty == 0) throw new Exception("訂單數量欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.OrderDate == null || item.OrderDate == "") throw new Exception("接單日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.PlannedShipmentDate == null || item.PlannedShipmentDate == "") throw new Exception("預計出貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);


                                    rowIndex++;

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CustomerMtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.OrderDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.SoQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.PlannedShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.PlannedShipmentQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ActualShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ActualShipmentQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Remark.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ShipFrom.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.PMDItem01.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PMDItem02.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.PMDItem03.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.PMDItem04.ToString().Trim();


                                    //rowIndex++;

                                }

                                #endregion

                                #region //設定

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion



                                #endregion

                                #region 輸出檔案
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    workbook.SaveAs(ms);
                                    excelBytes = ms.ToArray();
                                }
                                #endregion

                                #region//Mail 寄送托外進貨週期性報表

                                //取得Mail樣板設定
                                string SettingSchema = "SendPMDOrderExcel2";
                                string SettingNo = "Y";

                                var saleOrderDA = new SaleOrderDA();
                                string mailSettingsJson = saleOrderDA.SendCustomerMail(customerId);
                                if (string.IsNullOrEmpty(mailSettingsJson))
                                {
                                    throw new Exception("無法取得郵件設定");
                                }

                                var parsedSettings = JObject.Parse(mailSettingsJson);
                                if (parsedSettings["status"].ToString() != "success")
                                {
                                    throw new Exception(parsedSettings["msg"].ToString());
                                }

                                dynamic mailSettings = parsedSettings["data"];

                                foreach (var item in mailSettings)
                                {
                                    var mailFile = new MailFile
                                    {
                                        FileName = Path.GetFileNameWithoutExtension(fileName),
                                        FileExtension = ".xlsx",
                                        FileContent = excelBytes
                                    };


                                    #region //寄送Mail
                                    var mailConfig = new MailConfig
                                    {
                                        Host = item.Host,
                                        Port = Convert.ToInt32(item.Port),
                                        SendMode = Convert.ToInt32(item.SendMode),
                                        From = item.MailFrom,
                                        Subject = $"精定位订单明细_{DateTime.Now:yyyy/MM/dd}",
                                        Account = item.Account,
                                        Password = item.Password,
                                        MailTo = item.SalesMail,
                                        MailCc = item.SalesMailCc,
                                        MailBcc = item.SalesMailBcc,
                                        HtmlBody = $"附件為精定位订单明细，日期:{DateTime.Now:yyyy/MM/dd}",
                                        TextBody = "-",
                                        FileInfo = new List<MailFile> { mailFile },
                                        QcFileFlag = "N"
                                    };
                                    BaseHelper.MailSend(mailConfig);
                                    #endregion
                                }
                                #endregion

                                UpdatePageSended(pmdOrderIds, pmdOrdertempIds);

                                #region //Response
                                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "寄送完成"
                                });
                                #endregion

                            }
                            else
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "error",
                                    msg = "該客戶無資料"
                                });
                                #endregion

                                logger.Error("該客戶無資料");

                            }
                        }
                        #endregion



                    }
                    else
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = JObject.Parse(dataRequest)["msg"].ToString()
                        });
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                PMDAlertMamo(Company, "精定位订单明细", e.Message, "", "");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//SendPMDOrderExcel 寄送PMD-3
        [HttpPost]
        [Route("api/SCM/SendPMDOrderExcel3")]
        public void SendPMDOrderExcel3(string SendCustomer = "", string Company = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "SendPMDOrderExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion

                string[] sendCustomerList = SendCustomer.Split(',');
                foreach (var customer in sendCustomerList)
                {
                    int customerId = int.Parse(customer);

                    #region //Request
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetPMDOrderDataSend(3, customerId, Company);

                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                    {

                        byte[] excelBytes;
                        string fileName = $"粗胚模架-{DateTime.Now:yyyyMMdd}.xlsx";

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
                        defaultStyle.Font.FontName = "微軟正黑體";
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
                        titleStyle.Font.FontName = "微軟正黑體";
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
                        headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                        headerStyle.Font.FontName = "微軟正黑體";
                        headerStyle.Font.FontSize = 14;
                        headerStyle.Font.Bold = true;
                        #endregion

                        #region //dataStyle
                        var dataStyle = XLWorkbook.DefaultStyle;
                        dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.TopBorderColor = XLColor.Black;
                        dataStyle.Border.BottomBorderColor = XLColor.Black;
                        dataStyle.Border.LeftBorderColor = XLColor.Black;
                        dataStyle.Border.RightBorderColor = XLColor.Black;
                        dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                        dataStyle.Font.FontName = "微軟正黑體";
                        dataStyle.Font.FontSize = 12;
                        dataStyle.Font.Bold = false;
                        #endregion

                        #region //numberStyle
                        var numberStyle = XLWorkbook.DefaultStyle;
                        numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
                        numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                        numberStyle.Font.FontName = "微軟正黑體";
                        numberStyle.Font.FontSize = 12;
                        numberStyle.Font.Bold = false;
                        numberStyle.NumberFormat.Format = "#,##0.00";
                        #endregion

                        #region //currencyStyle
                        var currencyStyle = XLWorkbook.DefaultStyle;
                        currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "微軟正黑體";
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.NumberFormat.Format = "#,##0.00";
                        #endregion
                        #endregion

                        #region //EXCEL

                        using (var workbook = new XLWorkbook())
                        {

                            var worksheet = workbook.Worksheets.Add("粗胚模架");
                            worksheet.Style = defaultStyle;

                            // 設定標題
                            worksheet.Cell("A1").Value = "粗胚模架";
                            var titleRange = worksheet.Range("A1:R1");
                            titleRange.Merge().Style = titleStyle;
                            worksheet.Row(1).Height = 40;

                            // 設定表頭
                            string[] headers = new string[] { "廠內品號", "模號/品號", "品名", "QTY", "CAV", "接單日期", "預計出貨日", "實際出貨日", "實際出貨數量", "模架管理號", "備註"
                            , "發貨地", "半成品發貨(ZY->JC)", "預留欄位1", "預留欄位2", "預留欄位3", "預留欄位4" };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = worksheet.Cell(2, i + 1);
                                cell.Value = headers[i];
                                cell.Style = headerStyle;
                            }

                            // 寫入數據
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                            int rowIndex = 2;
                            string pmdOrderIds = string.Join(",", data.Select(d => d["PMDOrderId"].ToString()));
                            string pmdOrdertempIds = string.Join(",", data.Select(d => d["PMDOrderTempId"].ToString()));


                            #region //BODY
                            if (data.Count() > 0)
                            {
                                foreach (var item in data)
                                {
                                    if (item.CustomerId == null || item.CustomerId == 0) throw new Exception("客戶欄位不得為空! 資料序號: " + item.PMDOrderTempId);

                                    if (item.MtlItemNo == null || item.MtlItemNo == "") throw new Exception("場內品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.CustomerMtlItemNo == null || item.CustomerMtlItemNo == "") throw new Exception("模號/品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.MtlItemName == null || item.MtlItemNo == "") throw new Exception("品名欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.SoQty == null || item.SoQty == 0) throw new Exception("QTY欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.OrderDate == null || item.OrderDate == "") throw new Exception("接單日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.PlannedShipmentDate == null || item.PlannedShipmentDate == "") throw new Exception("預計出貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ActualShipmentDate == null || item.ActualShipmentDate == "") throw new Exception("實際出貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ShipFrom == null || item.ShipFrom == "") throw new Exception("發貨地欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    rowIndex++;

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CustomerMtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemName.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.SoQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.CAV.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.OrderDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.PlannedShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ActualShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.ActualShipmentQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.MoldFrameCtrlNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Remark.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.ShipFrom.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.SemiDelivery.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.PMDItem01.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.PMDItem02.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.PMDItem03.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).Value = item.PMDItem04.ToString().Trim();


                                    //rowIndex++;

                                }

                                #endregion

                                #region //設定

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion



                                #endregion

                                #region 輸出檔案
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    workbook.SaveAs(ms);
                                    excelBytes = ms.ToArray();
                                }
                                #endregion

                                #region//Mail 寄送托外進貨週期性報表

                                //取得Mail樣板設定
                                string SettingSchema = "SendPMDOrderExcel3";
                                string SettingNo = "Y";

                                var saleOrderDA = new SaleOrderDA();
                                string mailSettingsJson = saleOrderDA.SendCustomerMail(customerId);
                                if (string.IsNullOrEmpty(mailSettingsJson))
                                {
                                    throw new Exception("無法取得郵件設定");
                                }

                                var parsedSettings = JObject.Parse(mailSettingsJson);
                                if (parsedSettings["status"].ToString() != "success")
                                {
                                    throw new Exception(parsedSettings["msg"].ToString());
                                }

                                dynamic mailSettings = parsedSettings["data"];

                                foreach (var item in mailSettings)
                                {
                                    var mailFile = new MailFile
                                    {
                                        FileName = Path.GetFileNameWithoutExtension(fileName),
                                        FileExtension = ".xlsx",
                                        FileContent = excelBytes
                                    };


                                    #region //寄送Mail
                                    var mailConfig = new MailConfig
                                    {
                                        Host = item.Host,
                                        Port = Convert.ToInt32(item.Port),
                                        SendMode = Convert.ToInt32(item.SendMode),
                                        From = item.MailFrom,
                                        Subject = $"粗胚模架_{DateTime.Now:yyyy/MM/dd}",
                                        Account = item.Account,
                                        Password = item.Password,
                                        MailTo = item.SalesMail,
                                        MailCc = item.SalesMailCc,
                                        MailBcc = item.SalesMailBcc,
                                        HtmlBody = $"附件為粗胚模架，日期:{DateTime.Now:yyyy/MM/dd}",
                                        TextBody = "-",
                                        FileInfo = new List<MailFile> { mailFile },
                                        QcFileFlag = "N"
                                    };
                                    BaseHelper.MailSend(mailConfig);
                                    #endregion
                                }
                                #endregion

                                UpdatePageSended(pmdOrderIds, pmdOrdertempIds);

                                #region //Response
                                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "寄送完成"
                                });
                                #endregion


                            }
                            else
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "error",
                                    msg = "該客戶無資料"
                                });
                                #endregion

                                logger.Error("該客戶無資料");

                            }
                        }
                        #endregion



                    }
                    else
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = JObject.Parse(dataRequest)["msg"].ToString()
                        });
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                PMDAlertMamo(Company, "粗胚模架", e.Message, "", "");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//SendPMDOrderExcel 寄送PMD-4
        [HttpPost]
        [Route("api/SCM/SendPMDOrderExcel4")]
        public void SendPMDOrderExcel4(string SendCustomer = "", string Company = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "SendPMDOrderExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion

                string[] sendCustomerList = SendCustomer.Split(',');
                foreach (var customer in sendCustomerList)
                {
                    int customerId = int.Parse(customer);

                    #region //Request
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetPMDOrderDataSend(4, customerId, Company);

                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                    {

                        byte[] excelBytes;
                        string fileName = $"退货模芯衬装-{DateTime.Now:yyyyMMdd}.xlsx";

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
                        defaultStyle.Font.FontName = "微軟正黑體";
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
                        titleStyle.Font.FontName = "微軟正黑體";
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
                        headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                        headerStyle.Font.FontName = "微軟正黑體";
                        headerStyle.Font.FontSize = 14;
                        headerStyle.Font.Bold = true;
                        #endregion

                        #region //dataStyle
                        var dataStyle = XLWorkbook.DefaultStyle;
                        dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.TopBorderColor = XLColor.Black;
                        dataStyle.Border.BottomBorderColor = XLColor.Black;
                        dataStyle.Border.LeftBorderColor = XLColor.Black;
                        dataStyle.Border.RightBorderColor = XLColor.Black;
                        dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                        dataStyle.Font.FontName = "微軟正黑體";
                        dataStyle.Font.FontSize = 12;
                        dataStyle.Font.Bold = false;
                        #endregion

                        #region //numberStyle
                        var numberStyle = XLWorkbook.DefaultStyle;
                        numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
                        numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                        numberStyle.Font.FontName = "微軟正黑體";
                        numberStyle.Font.FontSize = 12;
                        numberStyle.Font.Bold = false;
                        numberStyle.NumberFormat.Format = "#,##0.00";
                        #endregion

                        #region //currencyStyle
                        var currencyStyle = XLWorkbook.DefaultStyle;
                        currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "微軟正黑體";
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.NumberFormat.Format = "#,##0.00";
                        #endregion
                        #endregion

                        #region //EXCEL

                        using (var workbook = new XLWorkbook())
                        {

                            var worksheet = workbook.Worksheets.Add("退货模芯衬装");
                            worksheet.Style = defaultStyle;

                            // 設定標題
                            worksheet.Cell("A1").Value = "退货模芯衬装";
                            var titleRange = worksheet.Range("A1:Q1");
                            titleRange.Merge().Style = titleStyle;
                            worksheet.Row(1).Height = 40;

                            // 設定表頭
                            string[] headers = new string[] { "退貨日期", "模號/品號", "品名", "退貨數量", "退貨問題點", "前期退貨數量", "前期欠數數量", "共需補貨數量", "預計出貨日", "實際出貨日", "備註"
                            , "發貨地", "預留欄位1", "預留欄位2", "預留欄位3", "預留欄位4" };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = worksheet.Cell(2, i + 1);
                                cell.Value = headers[i];
                                cell.Style = headerStyle;
                            }

                            // 寫入數據
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                            int rowIndex = 2;
                            string pmdOrderIds = string.Join(",", data.Select(d => d["PMDOrderId"].ToString()));
                            string pmdOrdertempIds = string.Join(",", data.Select(d => d["PMDOrderTempId"].ToString()));

                            #region //BODY
                            if (data.Count() > 0)
                            {
                                foreach (var item in data)
                                {
                                    if (item.CustomerId == null || item.CustomerId == 0) throw new Exception("客戶欄位不得為空! 資料序號: " + item.PMDOrderTempId);

                                    if (item.ReturnDate == null || item.ReturnDate == "") throw new Exception("退貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.CustomerMtlItemNo == null || item.CustomerMtlItemNo == "") throw new Exception("模號/品號欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.MtlItemName == null || item.MtlItemNo == "") throw new Exception("品名欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ReturnQty == null || item.ReturnQty == 0) throw new Exception("退貨數量欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ReturnIssue == null || item.ReturnIssue == "") throw new Exception("退貨問題點欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.PlannedShipmentDate == null || item.PlannedShipmentDate == "") throw new Exception("預計出貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ShipFrom == null || item.ShipFrom == "") throw new Exception("發貨地欄位不得為空! 資料序號: " + item.PMDOrderTempId);

                                    rowIndex++;

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ReturnDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.CustomerMtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ReturnQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.ReturnIssue.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.PrevReturnQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.PrevShortageQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.TotalReplenishmentQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.PlannedShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ActualShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.Remark.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.ShipFrom.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.PMDItem01.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.PMDItem02.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).Value = item.PMDItem03.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).Value = item.PMDItem04.ToString().Trim();


                                    //rowIndex++;

                                }

                                #endregion

                                #region //設定

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion



                                #endregion

                                #region 輸出檔案
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    workbook.SaveAs(ms);
                                    excelBytes = ms.ToArray();
                                }
                                #endregion

                                #region//Mail 寄送托外進貨週期性報表

                                //取得Mail樣板設定
                                string SettingSchema = "SendPMDOrderExcel4";
                                string SettingNo = "Y";

                                var saleOrderDA = new SaleOrderDA();
                                string mailSettingsJson = saleOrderDA.SendCustomerMail(customerId);
                                if (string.IsNullOrEmpty(mailSettingsJson))
                                {
                                    throw new Exception("無法取得郵件設定");
                                }

                                var parsedSettings = JObject.Parse(mailSettingsJson);
                                if (parsedSettings["status"].ToString() != "success")
                                {
                                    throw new Exception(parsedSettings["msg"].ToString());
                                }

                                dynamic mailSettings = parsedSettings["data"];

                                foreach (var item in mailSettings)
                                {
                                    var mailFile = new MailFile
                                    {
                                        FileName = Path.GetFileNameWithoutExtension(fileName),
                                        FileExtension = ".xlsx",
                                        FileContent = excelBytes
                                    };


                                    #region //寄送Mail
                                    var mailConfig = new MailConfig
                                    {
                                        Host = item.Host,
                                        Port = Convert.ToInt32(item.Port),
                                        SendMode = Convert.ToInt32(item.SendMode),
                                        From = item.MailFrom,
                                        Subject = $"退货模芯衬装_{DateTime.Now:yyyy/MM/dd}",
                                        Account = item.Account,
                                        Password = item.Password,
                                        MailTo = item.SalesMail,
                                        MailCc = item.SalesMailCc,
                                        MailBcc = item.SalesMailBcc,
                                        HtmlBody = $"附件為退货模芯衬装，日期:{DateTime.Now:yyyy/MM/dd}",
                                        TextBody = "-",
                                        FileInfo = new List<MailFile> { mailFile },
                                        QcFileFlag = "N"
                                    };
                                    BaseHelper.MailSend(mailConfig);
                                    #endregion
                                }
                                #endregion

                                UpdatePageSended(pmdOrderIds, pmdOrdertempIds);

                                #region //Response
                                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "寄送完成"
                                });
                                #endregion


                            }
                            else
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "error",
                                    msg = "該客戶無資料"
                                });
                                #endregion

                                logger.Error("該客戶無資料");

                            }
                        }
                        #endregion


                    }
                    else
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = JObject.Parse(dataRequest)["msg"].ToString()
                        });
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {

                PMDAlertMamo(Company, "退货模芯衬装", e.Message, "", "");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//SendPMDOrderExcel 寄送PMD-5
        [HttpPost]
        [Route("api/SCM/SendPMDOrderExcel5")]
        public void SendPMDOrderExcel5(string SendCustomer = "", string Company = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "SendPMDOrderExcel");
                #endregion

                #region //設定日期參數
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string StartReceiptDate = today;
                string EndReceiptDate = DateTime.Now.AddDays(-180).ToString("yyyy-MM-dd");

                #endregion

                string[] sendCustomerList = SendCustomer.Split(',');
                foreach (var customer in sendCustomerList)
                {
                    int customerId = int.Parse(customer);

                    #region //Request
                    saleOrderDA = new SaleOrderDA();
                    dataRequest = saleOrderDA.GetPMDOrderDataSend(5, customerId, Company);

                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                    {

                        byte[] excelBytes;
                        string fileName = $"精定位退货明细-{DateTime.Now:yyyyMMdd}.xlsx";

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
                        defaultStyle.Font.FontName = "微軟正黑體";
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
                        titleStyle.Font.FontName = "微軟正黑體";
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
                        headerStyle.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
                        headerStyle.Font.FontName = "微軟正黑體";
                        headerStyle.Font.FontSize = 14;
                        headerStyle.Font.Bold = true;
                        #endregion

                        #region //dataStyle
                        var dataStyle = XLWorkbook.DefaultStyle;
                        dataStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        dataStyle.Border.TopBorderColor = XLColor.Black;
                        dataStyle.Border.BottomBorderColor = XLColor.Black;
                        dataStyle.Border.LeftBorderColor = XLColor.Black;
                        dataStyle.Border.RightBorderColor = XLColor.Black;
                        dataStyle.Fill.BackgroundColor = XLColor.NoColor;
                        dataStyle.Font.FontName = "微軟正黑體";
                        dataStyle.Font.FontSize = 12;
                        dataStyle.Font.Bold = false;
                        #endregion

                        #region //numberStyle
                        var numberStyle = XLWorkbook.DefaultStyle;
                        numberStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        numberStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        numberStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        numberStyle.Border.TopBorderColor = XLColor.Black;
                        numberStyle.Border.BottomBorderColor = XLColor.Black;
                        numberStyle.Border.LeftBorderColor = XLColor.Black;
                        numberStyle.Border.RightBorderColor = XLColor.Black;
                        numberStyle.Fill.BackgroundColor = XLColor.NoColor;
                        numberStyle.Font.FontName = "微軟正黑體";
                        numberStyle.Font.FontSize = 12;
                        numberStyle.Font.Bold = false;
                        numberStyle.NumberFormat.Format = "#,##0.00";
                        #endregion

                        #region //currencyStyle
                        var currencyStyle = XLWorkbook.DefaultStyle;
                        currencyStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        currencyStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        currencyStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        currencyStyle.Border.TopBorderColor = XLColor.Black;
                        currencyStyle.Border.BottomBorderColor = XLColor.Black;
                        currencyStyle.Border.LeftBorderColor = XLColor.Black;
                        currencyStyle.Border.RightBorderColor = XLColor.Black;
                        currencyStyle.Fill.BackgroundColor = XLColor.NoColor;
                        currencyStyle.Font.FontName = "微軟正黑體";
                        currencyStyle.Font.FontSize = 12;
                        currencyStyle.Font.Bold = false;
                        currencyStyle.NumberFormat.Format = "#,##0.00";
                        #endregion
                        #endregion

                        #region //EXCEL

                        using (var workbook = new XLWorkbook())
                        {

                            var worksheet = workbook.Worksheets.Add("精定位退货明细");
                            worksheet.Style = defaultStyle;

                            // 設定標題
                            worksheet.Cell("A1").Value = "精定位退货明细";
                            var titleRange = worksheet.Range("A1:O1");
                            titleRange.Merge().Style = titleStyle;
                            worksheet.Row(1).Height = 40;

                            // 設定表頭
                            string[] headers = new string[] { "退貨日期", "退貨數量", "品名", "退貨問題點", "預計出貨日", "預計出貨數量", "實際出貨日", "實際出貨數量", "備註", "發貨地"
                            ,"預留欄位1", "預留欄位2", "預留欄位3", "預留欄位4" };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = worksheet.Cell(2, i + 1);
                                cell.Value = headers[i];
                                cell.Style = headerStyle;
                            }

                            // 寫入數據
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                            int rowIndex = 2;
                            string pmdOrderIds = string.Join(",", data.Select(d => d["PMDOrderId"].ToString()));
                            string pmdOrdertempIds = string.Join(",", data.Select(d => d["PMDOrderTempId"].ToString()));

                            #region //BODY
                            if (data.Count() > 0)
                            {
                                foreach (var item in data)
                                {

                                    if (item.CustomerId == null || item.CustomerId == 0) throw new Exception("客戶欄位不得為空! 資料序號: " + item.PMDOrderTempId);

                                    if (item.PlannedShipmentDate == null || item.PlannedShipmentDate == "") throw new Exception("預計出貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ActualShipmentDate == null || item.ActualShipmentDate == "") throw new Exception("實際出貨日期欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.PlannedShipmentQty == null || item.PlannedShipmentQty == 0) throw new Exception("預計出貨數量欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ActualShipmentQty == null || item.ActualShipmentQty == 0) throw new Exception("實際出貨數量欄位不得為空! 資料序號: " + item.PMDOrderTempId);
                                    if (item.ShipFrom == null || item.ShipFrom == "") throw new Exception("發貨地欄位不得為空! 資料序號: " + item.PMDOrderTempId);


                                    rowIndex++;

                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.ReturnDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.ReturnQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.CustomerMtlItemNo.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.ReturnIssue.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.PlannedShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.PlannedShipmentQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.ActualShipmentDate.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.ActualShipmentQty.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.Remark.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.ShipFrom.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.PMDItem01.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PMDItem02.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).Value = item.PMDItem03.ToString().Trim();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex)).Value = item.PMDItem04.ToString().Trim();
                                    //rowIndex++;

                                }

                                #endregion

                                #region //設定

                                #region //自適應欄寬
                                worksheet.Columns().AdjustToContents();
                                #endregion



                                #endregion

                                #region 輸出檔案
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    workbook.SaveAs(ms);
                                    excelBytes = ms.ToArray();
                                }
                                #endregion

                                #region//Mail 寄送托外進貨週期性報表

                                //取得Mail樣板設定
                                string SettingSchema = "SendPMDOrderExcel5";
                                string SettingNo = "Y";

                                var saleOrderDA = new SaleOrderDA();
                                string mailSettingsJson = saleOrderDA.SendCustomerMail(customerId);
                                if (string.IsNullOrEmpty(mailSettingsJson))
                                {
                                    throw new Exception("無法取得郵件設定");
                                }

                                var parsedSettings = JObject.Parse(mailSettingsJson);
                                if (parsedSettings["status"].ToString() != "success")
                                {
                                    throw new Exception(parsedSettings["msg"].ToString());
                                }

                                dynamic mailSettings = parsedSettings["data"];

                                foreach (var item in mailSettings)
                                {
                                    var mailFile = new MailFile
                                    {
                                        FileName = Path.GetFileNameWithoutExtension(fileName),
                                        FileExtension = ".xlsx",
                                        FileContent = excelBytes
                                    };


                                    #region //寄送Mail
                                    var mailConfig = new MailConfig
                                    {
                                        Host = item.Host,
                                        Port = Convert.ToInt32(item.Port),
                                        SendMode = Convert.ToInt32(item.SendMode),
                                        From = item.MailFrom,
                                        Subject = $"精定位退货明细{DateTime.Now:yyyy/MM/dd}",
                                        Account = item.Account,
                                        Password = item.Password,
                                        MailTo = item.SalesMail,
                                        MailCc = item.SalesMailCc,
                                        MailBcc = item.SalesMailBcc,
                                        HtmlBody = $"附件為精定位退货明细，日期:{DateTime.Now:yyyy/MM/dd}",
                                        TextBody = "-",
                                        FileInfo = new List<MailFile> { mailFile },
                                        QcFileFlag = "N"
                                    };
                                    BaseHelper.MailSend(mailConfig);
                                    #endregion
                                }
                                #endregion
                                UpdatePageSended(pmdOrderIds, pmdOrdertempIds);

                                #region //Response
                                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "寄送完成"
                                });
                                #endregion


                            }
                            else
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "error",
                                    msg = "該客戶無資料"
                                });
                                #endregion

                                logger.Error("該客戶無資料");

                            }
                        }
                        #endregion

                    }
                    else
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "error",
                            msg = JObject.Parse(dataRequest)["msg"].ToString()
                        });
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                PMDAlertMamo(Company, "精定位退货明细", e.Message, "", "");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //寄送舜宇PMD報表

        #region //SendPMDOrderExcelsSY 寄送舜宇PMD
        [HttpPost]
        [Route("api/SCM/SendPMDOrderExcelsSY")]
        public void SendPMDOrderExcelsSY()
        {
            SendPMDOrderExcel1("SY001");
            SendPMDOrderExcel2("SY001");
            SendPMDOrderExcel3("SY001");
            SendPMDOrderExcel4("SY001");
            SendPMDOrderExcel5("SY001");
        }
        #endregion
        #endregion


        #endregion

        #region //Download
        #region //Excel
        #region //SaleOrderExcelDownload 訂單資料匯出Excel
        public void SaleOrderExcelDownload(int SoId = -1, string SoErpPrefix = "", string SoErpNo = "", int CustomerId = -1
            , int SalesmenId = -1, string SearchKey = "", string ConfirmStatus = "", string ClosureStatus = ""
            , string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "excel");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrder02(SoId, SoErpPrefix, SoErpNo, CustomerId
                    , SalesmenId, SearchKey, ConfirmStatus, ClosureStatus
                    , StartDate, EndDate
                    , "", -1, -1);
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
                    string excelFileName = "訂單資料Excel檔";
                    string excelsheetName = "訂單詳細資料";

                    dynamic json = JsonConvert.DeserializeObject<ExpandoObject>(dataRequest, new ExpandoObjectConverter());
                    string[] header = new string[] { "單號", "單據日期", "客戶", "業務", "流水號", "品號", "品名", "數量", "單價", "金額", "預計出貨日", "備註" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 6)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 12).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        int startIndex = 0;
                        foreach (var item in json.data)
                        {
                            startIndex = rowIndex + 1;

                            if (item.SoDetail != null)
                            {
                                dynamic soDetail = JsonConvert.DeserializeObject<ExpandoObject>(item.SoDetail, new ExpandoObjectConverter());

                                foreach (var detail in soDetail.data)
                                {
                                    rowIndex++;
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.SoErpFullNo.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DocDate.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.CustomerNo.ToString() + " " + item.CustomerShortName.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.SalesmenNo.ToString() + " " + item.SalesmenName.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = "'" + detail.SoSequence.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = detail.MtlItemNo.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = detail.SoMtlItemName.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = detail.SoQty.ToString("N0");
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = detail.UnitPrice.ToString("N0");
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = detail.Amount.ToString("N0");
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = detail.PromiseDate.ToString();
                                    worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = detail.SoDetailRemark.ToString();
                                }
                            }
                            else
                            {
                                rowIndex++;
                                worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.SoErpFullNo.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DocDate.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.CustomerNo.ToString() + " " + item.CustomerShortName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.SalesmenNo.ToString() + " " + item.SalesmenName.ToString();
                                worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = "";
                                worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = "";
                            }
                            worksheet.Range(startIndex, 1, rowIndex, 1).Merge();
                            worksheet.Range(startIndex, 2, rowIndex, 2).Merge();
                            worksheet.Range(startIndex, 3, rowIndex, 3).Merge();
                            worksheet.Range(startIndex, 4, rowIndex, 4).Merge();
                        }

                        #endregion

                        #region //設定

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

        #region //Pdf
        #region //SaleOrderPdfDownload 訂單資料匯出Pdf
        public void SaleOrderPdfDownload(int SoId = -1, string ShowMoneyStatus = "")
        {
            try
            {
                WebApiLoginCheck("SaleOrderManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrderPdf(SoId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string SoFullNo = "";
                string htmlText = "";

                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    #region //html

                    if (result.Count() <= 0) throw new SystemException("此訂單無法進行列印作業!");

                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/ERP/SaleOrder.html"));

                    SoFullNo = result[0]["SoFullNo"].ToString();

                    htmlText = htmlText.Replace("[JmoLogo]", Server.MapPath("~/PdfTemplate/ERP/Logo.jpg"));

                    htmlText = htmlText.Replace("[SoErpPrefixNo]", result[0]["SoErpPrefixNo"].ToString());
                    htmlText = htmlText.Replace("[SoErpPrefixName]", result[0]["SoErpPrefixName"].ToString());
                    htmlText = htmlText.Replace("[SoErpNo]", result[0]["SoErpNo"].ToString());
                    htmlText = htmlText.Replace("[CustomerNo]", result[0]["CustomerNo"].ToString());
                    htmlText = htmlText.Replace("[CustomerName]", result[0]["CustomerName"].ToString());
                    htmlText = htmlText.Replace("[Currency]", result[0]["Currency"].ToString());
                    htmlText = htmlText.Replace("[TelNo]", result[0]["TelNo"].ToString());
                    htmlText = htmlText.Replace("[ExchangeRate]", result[0]["ExchangeRate"].ToString());
                    htmlText = htmlText.Replace("[DocDate]", Convert.ToDateTime(result[0]["DocDate"]).ToString("yyyy/MM/dd"));
                    htmlText = htmlText.Replace("[FaxNo]", result[0]["FaxNo"].ToString());
                    htmlText = htmlText.Replace("[BusinessTaxRate]", result[0]["BusinessTaxRate"].ToString());
                    htmlText = htmlText.Replace("[ShippingSiteNo]", result[0]["ShippingSiteNo"].ToString());
                    htmlText = htmlText.Replace("[ShippingSiteName]", result[0]["ShippingSiteName"].ToString());
                    htmlText = htmlText.Replace("[GuiNumber]", result[0]["GuiNumber"].ToString());
                    htmlText = htmlText.Replace("[TaxName]", result[0]["TaxationName"].ToString());
                    htmlText = htmlText.Replace("[DepartmentNo]", result[0]["DepartmentNo"].ToString());
                    htmlText = htmlText.Replace("[DepartmentName]", result[0]["DepartmentName"].ToString());
                    htmlText = htmlText.Replace("[PaymentTermNo]", result[0]["PaymentTermNo"].ToString());
                    htmlText = htmlText.Replace("[PaymentTermName]", result[0]["PaymentTermName"].ToString());
                    htmlText = htmlText.Replace("[SalesmenNo]", result[0]["SalesmenNo"].ToString());
                    htmlText = htmlText.Replace("[SalesmenName]", result[0]["SalesmenName"].ToString());
                    htmlText = htmlText.Replace("[CustomerPurchaseOrder]", result[0]["CustomerPurchaseOrder"].ToString());
                    htmlText = htmlText.Replace("[CustomerAddress]", result[0]["CustomerAddress"].ToString());

                    string htmlTemplate = htmlText;

                    var test = result[0]["SoDetail"];

                    if (result[0]["SoDetail"].ToString().Length > 0)
                    {
                        var detail = JObject.Parse(result[0]["SoDetail"].ToString())["data"];
                        string htmlDetail = "";
                        int mod = detail.Count() % 6;
                        int page = detail.Count() / 6 + (mod != 0 ? 1 : 0);

                        for (int p = 1; p <= page; p++)
                        {

                            if (p > 1)
                            {
                                htmlDetail = "";
                                htmlText += htmlTemplate;
                            }

                            for (int i = (p - 1) * 6; i < (p != page ? p * 6 : detail.Count()); i++)
                            {
                                htmlDetail += @"<tr>
                                                    <td style='text-align: center; height:63px; vertical-align: top;'>[SoSequence]</td>
                                                    <td>
                                                        <p>[MtlItemNo]</p>
                                                        <p>[MtlItemName]</p>
                                                        <p>[MtlItemSpec]</p>
                                                    </td>
                                                    <td style='text-align: right; vertical-align: top;'>
                                                        <p>[SoQty]</p>
                                                        <p></p>
                                                        <p>0</p>
                                                    </td>
                                                    <td style='text-align: right; vertical-align: top;'>[UomNo]</td>
                                                    <td style='text-align: right; vertical-align: top;'>
                                                        <p>[UnitPrice]</p>
                                                        <p>[Amount]</p>
                                                    </td>
                                                    <td style='text-align: center;display: flex;justify-content:center;border: none;'>
                                                        <div style='text-align: right; vertical-align: top;'>
                                                            <p>[PromiseDate]</p>
                                                            <p>[InventoryNo]</p>
                                                        </div>
                                                    </td>
                                                    <td style='vertical-align: top;'>[SoDetailRemark]</td>
                                                </tr>";

                                htmlDetail = htmlDetail.Replace("[SoSequence]", detail[i]["SoSequence"].ToString());
                                htmlDetail = htmlDetail.Replace("[MtlItemNo]", detail[i]["MtlItemNo"].ToString());
                                htmlDetail = htmlDetail.Replace("[MtlItemName]", detail[i]["MtlItemName"].ToString());
                                htmlDetail = htmlDetail.Replace("[MtlItemSpec]", detail[i]["MtlItemSpec"].ToString());
                                htmlDetail = htmlDetail.Replace("[SoQty]", Convert.ToInt32(detail[i]["SoQty"]).ToString("n0"));
                                if (ShowMoneyStatus == "N")
                                {
                                    htmlDetail = htmlDetail.Replace("[UnitPrice]", "");
                                    htmlDetail = htmlDetail.Replace("[Amount]", "");
                                }
                                else
                                {
                                    htmlDetail = htmlDetail.Replace("[UnitPrice]", Convert.ToDouble(detail[i]["UnitPrice"]).ToString());
                                    htmlDetail = htmlDetail.Replace("[Amount]", Convert.ToDouble(detail[i]["Amount"]).ToString("n0"));
                                }
                                htmlDetail = htmlDetail.Replace("[UomNo]", detail[i]["UomNo"].ToString());
                                htmlDetail = htmlDetail.Replace("[PromiseDate]", detail[i]["PromiseDate"].ToString());
                                htmlDetail = htmlDetail.Replace("[InventoryNo]", detail[i]["InventoryNo"].ToString());
                                htmlDetail = htmlDetail.Replace("[SoDetailRemark]", detail[i]["SoDetailRemark"].ToString());
                            }

                            htmlText = htmlText.Replace("[totalpage]", page.ToString());
                            htmlText = htmlText.Replace("[page]", p.ToString());

                            if (p == page)
                            {
                                if (mod != 0)
                                {
                                    htmlDetail += @"<tr>
                                                      <td style='height:63px;'></td>
                                                      <td>
                                                          以下空白 / /
                                                      </td>
                                                      <td style='text-align: right;'></td>
                                                      <td></td>
                                                      <td></td>
                                                      <td></td>
                                                      <td></td>
                                                   </tr>";

                                    for (int i = 0; i < 5 - mod; i++)
                                    {
                                        htmlDetail += @"<tr>
                                                      <td style='height:63px;'></td>
                                                      <td></td>
                                                      <td></td>
                                                      <td></td>
                                                      <td></td>
                                                      <td></td>
                                                      <td></td>
                                                    </tr>";
                                    }
                                }

                                htmlText = htmlText.Replace("[TotalQty]", Convert.ToInt32(result[0]["TotalQty"]).ToString("n0") + ".00");
                                if (ShowMoneyStatus == "N")
                                {
                                    htmlText = htmlText.Replace("[SoAmount]", "");
                                    htmlText = htmlText.Replace("[TaxAmount]", "");
                                    htmlText = htmlText.Replace("[TotalAmount]", "");
                                }
                                else
                                {
                                    htmlText = htmlText.Replace("[SoAmount]", Convert.ToInt32(result[0]["SoAmount"]).ToString("n0"));
                                    htmlText = htmlText.Replace("[TaxAmount]", Convert.ToInt32(result[0]["TaxAmount"]).ToString("n0"));
                                    htmlText = htmlText.Replace("[TotalAmount]", Convert.ToInt32(result[0]["TotalAmount"]).ToString("n0"));
                                }

                                htmlText = htmlText.Replace("[Approved]", "核准：");
                                htmlText = htmlText.Replace("[Finance]", "財務：");
                                htmlText = htmlText.Replace("[Supervisor]", "單位主管：");
                                htmlText = htmlText.Replace("[MakeBy]", "製表：");
                                htmlText = htmlText.Replace("[NextPage]", "");
                            }
                            else
                            {
                                htmlText = htmlText.Replace("[TotalQty]", "************.");
                                htmlText = htmlText.Replace("[SoAmount]", "************.**");
                                htmlText = htmlText.Replace("[TaxAmount]", "************.**");
                                htmlText = htmlText.Replace("[TotalAmount]", "************.**");

                                htmlText = htmlText.Replace("[Approved]", "");
                                htmlText = htmlText.Replace("[Finance]", "");
                                htmlText = htmlText.Replace("[Supervisor]", "");
                                htmlText = htmlText.Replace("[MakeBy]", "");
                                htmlText = htmlText.Replace("[NextPage]", "接下頁 . . . ");
                            }

                            htmlText = htmlText.Replace("[Detail]", htmlDetail);
                        }
                    }
                    else
                    {
                        string htmlDetail = @"<tr>
                                                <td style='height:63px;'></td>
                                                <td>以下空白 / /</td>
                                                <td style='text-align: right;'></td>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                              </tr>";

                        for (int i = 0; i < 5; i++)
                        {
                            htmlDetail += @"<tr>
                                                <td style='height:63px;'></td>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                                <td></td>
                                            </tr>";
                        }

                        htmlText = htmlText.Replace("[TotalQty]", "0.00");
                        htmlText = htmlText.Replace("[SoAmount]", "0");
                        htmlText = htmlText.Replace("[TaxAmount]", "0");
                        htmlText = htmlText.Replace("[TotalAmount]", "0");

                        htmlText = htmlText.Replace("[Approved]", "核准：");
                        htmlText = htmlText.Replace("[Finance]", "財務：");
                        htmlText = htmlText.Replace("[Supervisor]", "單位主管：");
                        htmlText = htmlText.Replace("[MakeBy]", "製表：");
                        htmlText = htmlText.Replace("[NextPage]", "");

                        htmlText = htmlText.Replace("[Detail]", htmlDetail);
                    }

                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (iTextSharp.text.Document document = new iTextSharp.text.Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
                        {
                            PdfWriter pdfWriter = PdfWriter.GetInstance(document, output); //指定文件預設開檔時的縮放為100%
                            PdfDestination pdfDestination = new PdfDestination(PdfDestination.XYZ, 0, document.PageSize.Height, 1f); //開啟Document文件 

                            document.Open(); //開啟Document文件 

                            XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, document, input, null, Encoding.UTF8, new UnicodeFontFactory()); //使用XMLWorkerHelper把Html parse到PDF檔裡

                            PdfAction action = PdfAction.GotoLocalPage(1, pdfDestination, pdfWriter); //將pdfDest設定的資料寫到PDF檔

                            pdfWriter.SetOpenAction(action);
                        }
                    }

                    pdfFile = output.ToArray();
                }

                string fileGuid = Guid.NewGuid().ToString();
                Session[fileGuid] = pdfFile;
                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = SoFullNo + " 訂單資料",
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

        #region //Word 
        #region //SaleOrderDocDownload 訂單資料單據下載
        public void SaleOrderDocDownload(int SoId = -1, string NonPrice = "")
        {
            try
            {
                if (!Regex.IsMatch(NonPrice, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【單據金額選項】錯誤!");

                #region //Request
                saleOrderDA = new SaleOrderDA();
                dataRequest = saleOrderDA.GetSaleOrderDoc(SoId);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = "", depositFilePath = "", user = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("單據資料不完整!");

                    bool showPrice = true;
                    switch (NonPrice)
                    {
                        case "N":
                            showPrice = true;
                            break;
                        case "Y":
                            showPrice = false;
                            break;
                    }

                    #region //判斷公司別文件
                    int CurrentCompany = Convert.ToInt32(Session["UserCompany"]);

                    if (Session["CompanySwitch"] != null)
                    {
                        if (Session["CompanySwitch"].ToString() == "manual")
                        {
                            CurrentCompany = Convert.ToInt32(Session["CompanyId"]);
                        }
                    }

                    switch (CurrentCompany)
                    {
                        case 2: //中揚
                            wordFileName = "客戶訂單-{0}-{1}";
                            filePath = "~/WordTemplate/SCM/P-S001表05-01 客戶訂單.docx";
                            secondFilePath = "~/WordTemplate/SCM/P-S001表05-01 客戶訂單 換頁.docx";
                            depositFilePath = "~/WordTemplate/SCM/訂金分批.docx";
                            break;
                        case 4: //晶彩
                            wordFileName = "客戶訂單-{0}-{1}";
                            filePath = "~/WordTemplate/SCM/R-SD01表05 客戶訂單-晶彩.docx";
                            secondFilePath = "~/WordTemplate/SCM/R-SD01表05 客戶訂單-晶彩 換頁.docx";
                            depositFilePath = "~/WordTemplate/SCM/訂金分批.docx";
                            break;
                        default:
                            throw new SystemException("目前公司別未開放功能!");
                    }
                    #endregion

                    var coptc = result["data"];
                    var coptd = result["dataDetail"];
                    var copuc = result["dataDeposit"];
                    user = result["user"].ToString();

                    #region //產生Doc
                    totalPage = coptd.Count() / 6 + (coptd.Count() % 6 > 0 ? 1 : 0);

                    if (totalPage == 1)
                    {
                        #region //單頁
                        using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                        {
                            wordFileName = string.Format(wordFileName, coptc[0]["TC001"].ToString(), coptc[0]["TC002"].ToString());

                            #region //單頭
                            doc.ReplaceText("[MQ002]", coptc[0]["MQ002"].ToString());
                            doc.ReplaceText("[TC001Order]", coptc[0]["TC001Order"].ToString());
                            doc.ReplaceText("[TC002]", coptc[0]["TC002"].ToString());
                            doc.ReplaceText("[TC039Date]", coptc[0]["TC039Date"].ToString());
                            doc.ReplaceText("[TC007Fac]", coptc[0]["TC007Fac"].ToString());
                            doc.ReplaceText("[TC005Dep]", coptc[0]["TC005Dep"].ToString());
                            doc.ReplaceText("[TC006Clerk]", coptc[0]["TC006Clerk"].ToString());
                            doc.ReplaceText("[TC004Clt]", coptc[0]["TC004Clt"].ToString());
                            doc.ReplaceText("[MA006]", coptc[0]["MA006"].ToString());
                            doc.ReplaceText("[MA008]", coptc[0]["MA008"].ToString());
                            doc.ReplaceText("[MA010]", coptc[0]["MA010"].ToString());
                            doc.ReplaceText("[TC014]", coptc[0]["TC014"].ToString());
                            doc.ReplaceText("[TC012]", coptc[0]["TC012"].ToString());
                            doc.ReplaceText("[TC010]", coptc[0]["TC010"].ToString());
                            doc.ReplaceText("[TC008]", coptc[0]["TC008"].ToString());
                            doc.ReplaceText("[TC009Rate]", coptc[0]["TC009Rate"].ToString());
                            doc.ReplaceText("[TC041]", coptc[0]["TC041"].ToString());
                            doc.ReplaceText("[TC016]", coptc[0]["TC016"].ToString());
                            doc.ReplaceText("[TC015]", coptc[0]["TC015"].ToString());

                            doc.ReplaceText("[cp]", "1");
                            doc.ReplaceText("[tp]", "1");

                            if (showPrice)
                            {
                                doc.ReplaceText("[TC029Amount]", coptc[0]["TC029Amount"].ToString());
                                doc.ReplaceText("[TC030Tax]", coptc[0]["TC030Tax"].ToString());
                                doc.ReplaceText("[TotalAmount]", coptc[0]["TotalAmount"].ToString());
                            }
                            else
                            {
                                doc.ReplaceText("[TC029Amount]", "");
                                doc.ReplaceText("[TC030Tax]", "");
                                doc.ReplaceText("[TotalAmount]", "");
                            }

                            doc.ReplaceText("[TC031Qua]", coptc[0]["TC031Qua"].ToString());
                            doc.ReplaceText("[ConfirmUser]", (user + "_" + coptc[0]["TC039"].ToString()));
                            #endregion

                            #region //單身
                            if (coptd.Count() > 0)
                            {
                                for (int i = 0; i < coptd.Count(); i++)
                                {
                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), coptd[i]["Seq"].ToString());
                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), coptd[i]["MtlItemNo"].ToString());

                                    string tempMtlItemName = coptd[i]["MtlItemName"].ToString();
                                    if (tempMtlItemName.Length > 25)
                                    {
                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 25) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                    string tempMtlItemSpec = coptd[i]["MtlItemSpec"].ToString();
                                    if (tempMtlItemSpec.Length > 25)
                                    {
                                        tempMtlItemSpec = BaseHelper.StrLeft(tempMtlItemSpec, 25) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), tempMtlItemSpec);

                                    string tempCutMtlItemName = coptd[i]["CutMtlItemName"].ToString();
                                    if (tempCutMtlItemName.Length > 33)
                                    {
                                        tempCutMtlItemName = BaseHelper.StrLeft(tempCutMtlItemName, 33) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[CutMtlItemName{0:00}]", i + 1), tempCutMtlItemName);

                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), coptd[i]["Qty"].ToString());
                                    doc.ReplaceText(string.Format("[GQty{0:00}]", i + 1), coptd[i]["GQty"].ToString());
                                    doc.ReplaceText(string.Format("[PQty{0:00}]", i + 1), coptd[i]["PQty"].ToString());
                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), coptd[i]["Uom"].ToString());

                                    if (showPrice)
                                    {
                                        doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), coptd[i]["UnPri"].ToString());
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), coptd[i]["Amount"].ToString());
                                    }
                                    else
                                    {
                                        doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                    }

                                    doc.ReplaceText(string.Format("[PrDate{0:00}]", i + 1), coptd[i]["PrDate"].ToString());
                                    doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), coptd[i]["InvNo"].ToString());

                                    string tempRemark = coptd[i]["Remark"].ToString();
                                    if (tempRemark.Length > 50)
                                    {
                                        tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                    }
                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                }

                                #region //剩餘欄位
                                if (coptd.Count() < 6)
                                {
                                    for (int i = coptd.Count(); i < 6; i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                        if (i == coptd.Count() && i < 6)
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                        }
                                        else
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                        }

                                        doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[CutMtlItemName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[GQty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[PQty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[PrDate{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");
                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("【單身資料】錯誤!");
                            }
                            #endregion

                            #region //定金分批頁面
                            if (coptc[0]["TC070"].ToString() == "Y")
                            {
                                using (DocX depositDoc = DocX.Load(Server.MapPath(depositFilePath)))
                                {
                                    doc.InsertDocument(depositDoc, true);
                                }

                                doc.ReplaceText("[MQ002]", coptc[0]["MQ002"].ToString());
                                doc.ReplaceText("[TC001Order]", coptc[0]["TC001Order"].ToString());
                                doc.ReplaceText("[TC002]", coptc[0]["TC002"].ToString());

                                doc.ReplaceText("[cp]", "1");
                                doc.ReplaceText("[tp]", "1");

                                if (copuc.Count() > 0)
                                {
                                    for (int i = 0; i < copuc.Count(); i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), copuc[i]["Seq"].ToString());
                                        doc.ReplaceText(string.Format("[Ratio{0:00}]", i + 1), copuc[i]["Ratio"].ToString() + "%");
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), Convert.ToDouble(copuc[i]["Amount"]).ToString("n0"));
                                        doc.ReplaceText(string.Format("[PaymentDate{0:00}]", i + 1), copuc[i]["PaymentDate"].ToString());
                                        doc.ReplaceText(string.Format("[ClosingStatus{0:00}]", i + 1), copuc[i]["ClosingStatus"].ToString());
                                    }

                                    #region //剩餘欄位
                                    if (copuc.Count() < 10)
                                    {
                                        for (int i = copuc.Count(); i < 10; i++)
                                        {
                                            if (i == copuc.Count() && i < 10)
                                            {
                                                doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[Ratio{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[PaymentDate{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ClosingStatus{0:00}]", i + 1), "");
                                        }
                                    }
                                    #endregion
                                }
                            }
                            #endregion

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region //多頁
                        using (DocX doc = DocX.Load(Server.MapPath(secondFilePath)))
                        {
                            wordFileName = string.Format(wordFileName, coptc[0]["TC001"].ToString(), coptc[0]["TC002"].ToString());

                            for (int p = 1; p <= totalPage; p++)
                            {
                                if (p != totalPage)
                                {
                                    #region //預設頁面
                                    if (p != 1)
                                    {
                                        using (DocX subDoc = DocX.Load(Server.MapPath(secondFilePath)))
                                        {
                                            doc.InsertDocument(subDoc, true);
                                        }
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", coptc[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TC001Order]", coptc[0]["TC001Order"].ToString());
                                    doc.ReplaceText("[TC002]", coptc[0]["TC002"].ToString());
                                    doc.ReplaceText("[TC039Date]", coptc[0]["TC039Date"].ToString());
                                    doc.ReplaceText("[TC007Fac]", coptc[0]["TC007Fac"].ToString());
                                    doc.ReplaceText("[TC005Dep]", coptc[0]["TC005Dep"].ToString());
                                    doc.ReplaceText("[TC006Clerk]", coptc[0]["TC006Clerk"].ToString());
                                    doc.ReplaceText("[TC004Clt]", coptc[0]["TC004Clt"].ToString());
                                    doc.ReplaceText("[MA006]", coptc[0]["MA006"].ToString());
                                    doc.ReplaceText("[MA008]", coptc[0]["MA008"].ToString());
                                    doc.ReplaceText("[MA010]", coptc[0]["MA010"].ToString());
                                    doc.ReplaceText("[TC014]", coptc[0]["TC014"].ToString());
                                    doc.ReplaceText("[TC012]", coptc[0]["TC012"].ToString());
                                    doc.ReplaceText("[TC010]", coptc[0]["TC010"].ToString());
                                    doc.ReplaceText("[TC008]", coptc[0]["TC008"].ToString());
                                    doc.ReplaceText("[TC009Rate]", coptc[0]["TC009Rate"].ToString());
                                    doc.ReplaceText("[TC041]", coptc[0]["TC041"].ToString());
                                    doc.ReplaceText("[TC016]", coptc[0]["TC016"].ToString());
                                    doc.ReplaceText("[TC015]", coptc[0]["TC015"].ToString());


                                    doc.ReplaceText("[TC031Qua]", "************");
                                    doc.ReplaceText("[TC029Amount]", "************.**");
                                    doc.ReplaceText("[TC030Tax]", "************.**");
                                    doc.ReplaceText("[TotalAmount]", "************.**");
                                    doc.ReplaceText("[cp]", p.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[NextPage]", "接下頁...");
                                    #endregion

                                    #region //單身
                                    if (coptd.Count() > 0)
                                    {
                                        for (int i = 0; i < 6; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), coptd[i + (p - 1) * 6]["MtlItemNo"].ToString());

                                            string tempMtlItemName = coptd[i + (p - 1) * 6]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 25)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 25) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), coptd[i + (p - 1) * 6]["MtlItemSpec"].ToString());

                                            string tempCutMtlItemName = coptd[i + (p - 1) * 6]["CutMtlItemName"].ToString();
                                            if (tempCutMtlItemName.Length > 33)
                                            {
                                                tempCutMtlItemName = BaseHelper.StrLeft(tempCutMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[CutMtlItemName{0:00}]", i + 1), tempCutMtlItemName);

                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Qty"].ToString());
                                            doc.ReplaceText(string.Format("[GQty{0:00}]", i + 1), coptd[i + (p - 1) * 6]["GQty"].ToString());
                                            doc.ReplaceText(string.Format("[PQty{0:00}]", i + 1), coptd[i + (p - 1) * 6]["PQty"].ToString());
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Uom"].ToString());

                                            if (showPrice)
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), coptd[i + (p - 1) * 6]["UnPri"].ToString());
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Amount"].ToString());
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[PrDate{0:00}]", i + 1), coptd[i + (p - 1) * 6]["PrDate"].ToString());
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), coptd[i + (p - 1) * 6]["InvNo"].ToString());

                                            string tempRemark = coptd[i + (p - 1) * 6]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion
                                    #endregion
                                }
                                else
                                {
                                    #region //最後一頁
                                    using (DocX finalDoc = DocX.Load(Server.MapPath(filePath)))
                                    {
                                        doc.InsertDocument(finalDoc, true);
                                    }

                                    #region //單頭
                                    doc.ReplaceText("[MQ002]", coptc[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TC001Order]", coptc[0]["TC001Order"].ToString());
                                    doc.ReplaceText("[TC002]", coptc[0]["TC002"].ToString());
                                    doc.ReplaceText("[TC039Date]", coptc[0]["TC039Date"].ToString());
                                    doc.ReplaceText("[TC007Fac]", coptc[0]["TC007Fac"].ToString());
                                    doc.ReplaceText("[TC005Dep]", coptc[0]["TC005Dep"].ToString());
                                    doc.ReplaceText("[TC006Clerk]", coptc[0]["TC006Clerk"].ToString());
                                    doc.ReplaceText("[TC004Clt]", coptc[0]["TC004Clt"].ToString());
                                    doc.ReplaceText("[MA006]", coptc[0]["MA006"].ToString());
                                    doc.ReplaceText("[MA008]", coptc[0]["MA008"].ToString());
                                    doc.ReplaceText("[MA010]", coptc[0]["MA010"].ToString());
                                    doc.ReplaceText("[TC014]", coptc[0]["TC014"].ToString());
                                    doc.ReplaceText("[TC012]", coptc[0]["TC012"].ToString());
                                    doc.ReplaceText("[TC010]", coptc[0]["TC010"].ToString());
                                    doc.ReplaceText("[TC008]", coptc[0]["TC008"].ToString());
                                    doc.ReplaceText("[TC009Rate]", coptc[0]["TC009Rate"].ToString());
                                    doc.ReplaceText("[TC041]", coptc[0]["TC041"].ToString());
                                    doc.ReplaceText("[TC016]", coptc[0]["TC016"].ToString());
                                    doc.ReplaceText("[TC015]", coptc[0]["TC015"].ToString());

                                    doc.ReplaceText("[cp]", totalPage.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[TC031Qua]", coptc[0]["TC031Qua"].ToString());

                                    if (showPrice)
                                    {
                                        doc.ReplaceText("[TC029Amount]", coptc[0]["TC029Amount"].ToString());
                                        doc.ReplaceText("[TC030Tax]", coptc[0]["TC030Tax"].ToString());
                                        doc.ReplaceText("[TotalAmount]", coptc[0]["TotalAmount"].ToString());
                                    }
                                    else
                                    {
                                        doc.ReplaceText("[TC029Amount]", "");
                                        doc.ReplaceText("[TC030Tax]", "");
                                        doc.ReplaceText("[TotalAmount]", "");
                                    }

                                    doc.ReplaceText("[ConfirmUser]", user + "_" + coptc[0]["TC039"].ToString());
                                    #endregion

                                    #region //單身
                                    if (coptd.Count() > 0)
                                    {
                                        for (int i = 0; i < (coptd.Count() % 6 != 0 ? coptd.Count() % 6 : 6); i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), coptd[i + (p - 1) * 6]["MtlItemNo"].ToString());

                                            string tempMtlItemName = coptd[i + (p - 1) * 6]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 25)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 25) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), coptd[i + (p - 1) * 6]["MtlItemSpec"].ToString());

                                            string tempCutMtlItemName = coptd[i + (p - 1) * 6]["CutMtlItemName"].ToString();
                                            if (tempCutMtlItemName.Length > 33)
                                            {
                                                tempCutMtlItemName = BaseHelper.StrLeft(tempCutMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[CutMtlItemName{0:00}]", i + 1), tempCutMtlItemName);

                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Qty"].ToString());
                                            doc.ReplaceText(string.Format("[GQty{0:00}]", i + 1), coptd[i + (p - 1) * 6]["GQty"].ToString());
                                            doc.ReplaceText(string.Format("[PQty{0:00}]", i + 1), coptd[i + (p - 1) * 6]["PQty"].ToString());
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Uom"].ToString());

                                            if (showPrice)
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), coptd[i + (p - 1) * 6]["UnPri"].ToString());
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), coptd[i + (p - 1) * 6]["Amount"].ToString());
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[PrDate{0:00}]", i + 1), coptd[i + (p - 1) * 6]["PrDate"].ToString());
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), coptd[i + (p - 1) * 6]["InvNo"].ToString());

                                            string tempRemark = coptd[i + (p - 1) * 6]["Remark"].ToString();
                                            if (tempRemark.Length > 50)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 50) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                        }
                                    }
                                    else
                                    {
                                        throw new SystemException("【單身資料】錯誤!");
                                    }
                                    #endregion

                                    #region //剩餘欄位
                                    if (coptd.Count() % 6 != 0)
                                    {
                                        for (int i = coptd.Count() % 6; i < 6; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                            if (i == coptd.Count() % 6 && i < 6)
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[CutMtlItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[GQty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[PQty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");

                                            if (showPrice)
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[PrDate{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");
                                        }
                                    }
                                    #endregion
                                    #endregion
                                }
                            }

                            #region //定金分批頁面
                            if (coptc[0]["TC070"].ToString() == "Y")
                            {
                                using (DocX depositDoc = DocX.Load(Server.MapPath(depositFilePath)))
                                {
                                    doc.InsertDocument(depositDoc, true);
                                }

                                doc.ReplaceText("[MQ002]", coptc[0]["MQ002"].ToString());
                                doc.ReplaceText("[TC001Order]", coptc[0]["TC001Order"].ToString());
                                doc.ReplaceText("[TC002]", coptc[0]["TC002"].ToString());

                                doc.ReplaceText("[cp]", "1");
                                doc.ReplaceText("[tp]", "1");

                                if (copuc.Count() > 0)
                                {
                                    for (int i = 0; i < copuc.Count(); i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), copuc[i]["Seq"].ToString());
                                        doc.ReplaceText(string.Format("[Ratio{0:00}]", i + 1), copuc[i]["Ratio"].ToString() + "%");
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), Convert.ToDouble(copuc[i]["Amount"]).ToString("n0"));
                                        doc.ReplaceText(string.Format("[PaymentDate{0:00}]", i + 1), copuc[i]["PaymentDate"].ToString());
                                        doc.ReplaceText(string.Format("[ClosingStatus{0:00}]", i + 1), copuc[i]["ClosingStatus"].ToString());
                                    }

                                    #region //剩餘欄位
                                    if (copuc.Count() < 10)
                                    {
                                        for (int i = copuc.Count(); i < 10; i++)
                                        {
                                            if (i == copuc.Count() && i < 10)
                                            {
                                                doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[Ratio{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[PaymentDate{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ClosingStatus{0:00}]", i + 1), "");
                                        }
                                    }
                                    #endregion
                                }
                            }
                            #endregion

                            using (MemoryStream output = new MemoryStream())
                            {
                                doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                                doc.SaveAs(output, password);

                                wordFile = output.ToArray();
                            }
                        }
                        #endregion
                    }

                    fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = wordFile;
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = wordFileName,
                        fileExtension = ".docx"
                    });
                    #endregion
                }
                else
                {
                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "目前尚無資料"
                    });
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
        #endregion
    }
}