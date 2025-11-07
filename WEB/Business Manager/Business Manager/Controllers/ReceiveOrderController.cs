using Helpers;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Business_Manager.Controllers
{
    public class ReceiveOrderController : WebController
    {
        private ReceiveOrderDA receiveOrderDA = new ReceiveOrderDA();      

        #region//View
        public ActionResult ReceiveOrderManagment()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult ReturnReceiveOrderManagment()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region//Get
        #region //GetReceiveOrder 取得銷貨單單據資料
        [HttpPost]
        public void GetReceiveOrder(int RoId = -1, string RoErpPrefix= "", string RoErpFullNo = "", int CustomerId = -1, int SalesmenId = -1
            , string StartDocDate = "", string EndDocDate = "", string ConfirmStatus = "", string TransferStatusMES = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetReceiveOrder(RoId, RoErpPrefix, RoErpFullNo, CustomerId, SalesmenId
                    , StartDocDate, EndDocDate, ConfirmStatus, TransferStatusMES
                    , OrderBy, PageIndex, PageSize
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

        #region //GetRoDetail 取得銷貨單單身資料
        [HttpPost]
        public void GetRoDetail(int RoId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetRoDetail(RoId
                    , OrderBy, PageIndex, PageSize
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

        #region //GetRoErpPrefixSetting 取得銷貨單單別設定
        [HttpPost]
        public void GetRoErpPrefixSetting( string RoErpPrefix = "")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetRoErpPrefixSetting(RoErpPrefix);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetExchangeRate 取得ERP匯率資料
        [HttpPost]
        public void GetExchangeRate(string Condition = "", string ErpPrefix = "", string OrderBy = "")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetExchangeRate(Condition, ErpPrefix, OrderBy
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

        #region //GetCurrencyDouble 取得幣別小數取位
        [HttpPost]
        public void GetCurrencyDouble(string Currency = "")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetCurrencyDouble(Currency
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

        #region //GetInventory 取得ERP庫別資料 
        [HttpPost]
        public void GetInventory(int ViewCompanyId = -1, string Table = "", string InventoryNo = "")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetInventory(ViewCompanyId, Table, InventoryNo
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

        #region //GetSourceOrderData 取得來源單據資料
        [HttpPost]
        public void GetSourceOrderData(string SourceType = "", string SourcePrefix = "", string SourceNo = "", string SourceSequence = "", string DoDetailList ="")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetSourceOrderData(SourceType, SourcePrefix, SourceNo, SourceSequence, DoDetailList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMtlItemUom 取得品號單位
        [HttpPost]
        public void GetMtlItemUom(string MtlItemNo = "")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetMtlItemUom(MtlItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetReturnReceiveOrder 取得銷退單資料
        [HttpPost]
        public void GetReturnReceiveOrder(int RtId = -1, string RtErpPrefix = "", string RtErpFullNo = "", int CustomerId = -1, int SalesmenId = -1
            , string StartDocDate = "", string EndDocDate = "", string ConfirmStatus = "", string TransferStatusMES = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetReturnReceiveOrder(RtId, RtErpPrefix, RtErpFullNo, CustomerId, SalesmenId
                    , StartDocDate, EndDocDate, ConfirmStatus, TransferStatusMES
                    , OrderBy, PageIndex, PageSize
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

        #region //GetRtDetail 取得銷退單單身資料
        [HttpPost]
        public void GetRtDetail(int RtId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetRtDetail(RtId
                    , OrderBy, PageIndex, PageSize
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

        #region //GetDeliveryOrder 取得銷貨單單據資料
        [HttpPost]
        public void GetDeliveryOrder(int CustomerId = -1, string SearchKey = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "read");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.GetDeliveryOrder(CustomerId, SearchKey, StartDate, EndDate
                    , OrderBy, PageIndex, PageSize
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

        #endregion

        #region //Add
        #region //AddReceiveOrder 銷貨單新增
        [HttpPost]
        public void AddReceiveOrder(int ViewCompanyId = -1
            , string DocDate = ""
            , string ReceiveDate = "", string RoErpPrefix = "", string RoErpNo = "", int CustomerId = -1, string ProcessCode = ""
            , string RevenueJournalCode = "", string CostJournalCode = "", string NoCredLimitControl = "", string CashSales = ""
            , string SourceType = "", string PriceSourceTypeMain = "", int SourceOrderId = -1, string SourceFull = ""
            , int DepartmentId = -1
            , double OriginalAmount = -1, double OriginalTaxAmount = -1, double TotalQuantity = -1, double PretaxAmount = -1, double TaxAmount = -1
            , double TaxCurrencyRate = -1
            , int SalesmenId = -1, string Factory = "", string PaymentTerm = "", string TradingTerms = "", string Currency = ""
            , double ExchangeRate = -1, string BondCode = "", int RowCnt = -1, string Remark = ""
            , string CustomerFullName = "", string ContactPerson = "", string CustomerAddressFirst = "", string TelephoneNumber = ""
            , string CustomerAddressSecond = "", string FaxNumber = ""
            , string TaxCode = ""
            , string InvoiceType = "", string TaxType = "", double TaxRate = -1, string InvNumGenMethod = "", string UiNo = ""
            , string InvoiceDate = "", string InvoiceTime = "", string InvNumStart = "", string InvNumEnd = "", string ApplyYYMM = "", string CustomsClearance = ""
            , string CustomerFullNameOre = "", string InvoiceAddressFirst = "", string CustomerEgFullName = "", string InvoiceAddressSecond = ""
            , string MultipleInvoices = "", string MultiTaxRate = "", string AttachInvWithShip = "", string InvoicesVoid = ""
            , string VehicleTypeNumber = ""
            , string InvDonationRecipient = "", string VehicleIDshow = "", string VehicleIDhide = "", string CreditCard4No = "", string InvRandomCode = "", string ContactEmail = ""
            , int StaffId = -1
            , int CollectionSalesmenId = -1, string Remark1 = "", string Remark2 = "", string Remark3 = "", string LCNO = "", string INVOICENO = ""
            , string DeclarationNumber = "", string NewRoFull = "", string ShipNoticeFull = "", string ChangeInvCode = ""
            , string DepositBatches = ""
            , string SoFull = "", string AdvOrderFull = "", double OffsetAmount = -1, double OffsetTaxAmount = -1, string TransportMethod = ""
            , string DispatchOrderFull = "", string CarNumber = "", string TrainNumber = "", string DeliveryUser = ""
            , string Courier = "", string SiteCommCalcMethod = "", string SiteCommRate = "", string TotalCommAmount = ""
            , string RoDetailData = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "add");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.AddReceiveOrder(ViewCompanyId
                        , DocDate
                        , ReceiveDate, RoErpPrefix, RoErpNo, CustomerId, ProcessCode
                        , RevenueJournalCode, CostJournalCode, NoCredLimitControl, CashSales
                        , SourceType, PriceSourceTypeMain, SourceOrderId, SourceFull
                        , DepartmentId
                        , OriginalAmount, OriginalTaxAmount, TotalQuantity, PretaxAmount, TaxAmount
                        , TaxCurrencyRate
                        , SalesmenId, Factory, PaymentTerm, TradingTerms, Currency
                        , ExchangeRate, BondCode, RowCnt, Remark
                        , CustomerFullName, ContactPerson, CustomerAddressFirst, TelephoneNumber
                        , CustomerAddressSecond, FaxNumber
                        , TaxCode
                        , InvoiceType, TaxType, TaxRate, InvNumGenMethod, UiNo
                        , InvoiceDate, InvoiceTime, InvNumStart, InvNumEnd, ApplyYYMM, CustomsClearance
                        , CustomerFullNameOre, InvoiceAddressFirst, CustomerEgFullName, InvoiceAddressSecond
                        , MultipleInvoices, MultiTaxRate, AttachInvWithShip, InvoicesVoid
                        , VehicleTypeNumber
                        , InvDonationRecipient, VehicleIDshow, VehicleIDhide, CreditCard4No, InvRandomCode, ContactEmail
                        , StaffId
                        , CollectionSalesmenId, Remark1, Remark2, Remark3, LCNO, INVOICENO
                        , DeclarationNumber, NewRoFull, ShipNoticeFull, ChangeInvCode
                        , DepositBatches
                        , SoFull, AdvOrderFull, OffsetAmount, OffsetTaxAmount, TransportMethod
                        , DispatchOrderFull, CarNumber, TrainNumber, DeliveryUser
                        , Courier, SiteCommCalcMethod, SiteCommRate, TotalCommAmount
                        , RoDetailData
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

        #region //AddReturnReceiveOrder 銷退單新增
        [HttpPost]
        public void AddReturnReceiveOrder(int ViewCompanyId = -1
            , string DocDate = ""
            , string ReturnDate = "", string RtErpPrefix = "", string RtErpNo = "", int CustomerId = -1, string ProcessCode = ""
            , string RevenueJournalCode = "", string CostJournalCode = ""
            , double OriginalAmount = -1
            , double OriginalTaxAmount = -1, double PretaxAmount = -1, double TaxAmount = -1, double TotalQuantity = -1
            , double TaxCurrencyRate = -1
            , int DepartmentId = -1
            , int SalesmenId = -1, int StaffId = -1, int CollectionSalesmenId = -1, string Factory = "", string PaymentTerm = "", string TradingTerms = ""
            , string Currency = "", double ExchangeRate = -1, string BondCode = "", int RowCnt = -1, string Remark = ""
            , string TaxCode = ""
            , string InvoiceType = "", string TaxType = "", double TaxRate = -1, string UiNo = "", string InvoiceDate = "", string InvNumStart = "", string InvNumEnd = ""
            , string ApplyYYMM = "", string CustomsClearance = ""
            , string CustomerFullName = ""
            , string CustomerEgFullName = "", string MultipleInvoices = "", string MultiTaxRate = "", string DebitNote = ""
            , string ContactPerson = ""
            , string ContactEmail = "", string CustomerAddressFirst = "", string CustomerAddressSecond = "", string Remark1 = "", string Remark2 = "", string Remark3 = ""
            , string RtDetailData = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "add");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.AddReturnReceiveOrder(ViewCompanyId,
                            DocDate,
                            ReturnDate,RtErpPrefix,RtErpNo,CustomerId,ProcessCode,
                            RevenueJournalCode,CostJournalCode,
                            OriginalAmount,
                            OriginalTaxAmount,PretaxAmount,TaxAmount,TotalQuantity,
                            TaxCurrencyRate,
                            DepartmentId,
                            SalesmenId,StaffId,CollectionSalesmenId,Factory,PaymentTerm,TradingTerms,
                            Currency,ExchangeRate,BondCode,RowCnt,Remark,
                            TaxCode,
                            InvoiceType,TaxType,TaxRate,UiNo,InvoiceDate,InvNumStart,InvNumEnd,
                            ApplyYYMM,CustomsClearance,
                            CustomerFullName,CustomerEgFullName,MultipleInvoices,DebitNote,
                            ContactPerson,
                            ContactEmail,CustomerAddressFirst,CustomerAddressSecond,Remark1,Remark2,Remark3,
                            RtDetailData
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

        #region //AddReceiveOrderToErp 出貨單轉銷貨
        [HttpPost]
        public void AddReceiveOrderToErp(int ViewCompanyId = -1, string DocDate = "", string RoErpPrefix = ""
            , int CustomerId = -1, int SalesmenId = -1, int CollectionSalesmenId = -1, int DepartmentId = -1, string Remark = "", string InvNumStart = "", string DoDetails = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "add");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.AddReceiveOrderToErp(ViewCompanyId, DocDate, RoErpPrefix
                        , CustomerId, SalesmenId, CollectionSalesmenId, DepartmentId, Remark, InvNumStart, DoDetails
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


        #endregion

        #region //Update
        #region //UpdateReceiveOrder 銷貨單新增
        [HttpPost]
        public void UpdateReceiveOrder(int RoId, int ViewCompanyId = -1
            , string DocDate = ""
            , string ReceiveDate = "", int CustomerId = -1, string ProcessCode = ""
            , string RevenueJournalCode = "", string CostJournalCode = "", string NoCredLimitControl = "", string CashSales = ""
            , string SourceType = "", string PriceSourceTypeMain = "", int SourceOrderId =-1, string SourceFull = ""
            , int DepartmentId = -1
            , double OriginalAmount = -1, double OriginalTaxAmount = -1, double TotalQuantity = -1, double PretaxAmount = -1, double TaxAmount = -1
            , double TaxCurrencyRate = -1
            , int SalesmenId = -1, string Factory = "", string PaymentTerm = "", string TradingTerms = "", string Currency = ""
            , double ExchangeRate = -1, string BondCode = "", int RowCnt = -1, string Remark = ""
            , string CustomerFullName = "", string ContactPerson = "", string CustomerAddressFirst = "", string TelephoneNumber = ""
            , string CustomerAddressSecond = "", string FaxNumber = ""
            , string TaxCode = ""
            , string InvoiceType = "", string TaxType = "", double TaxRate = -1, string InvNumGenMethod = "", string UiNo = ""
            , string InvoiceDate = "", string InvoiceTime = "", string InvNumStart = "", string InvNumEnd = "", string ApplyYYMM = "", string CustomsClearance = ""
            , string CustomerFullNameOre = "", string InvoiceAddressFirst = "", string CustomerEgFullName = "", string InvoiceAddressSecond = ""
            , string MultipleInvoices = "", string MultiTaxRate = "", string AttachInvWithShip = "", string InvoicesVoid = ""
            , string VehicleTypeNumber = ""
            , string InvDonationRecipient = "", string VehicleIDshow = "", string VehicleIDhide = "", string CreditCard4No = "", string InvRandomCode = "", string ContactEmail = ""
            , int StaffId = -1
            , int CollectionSalesmenId = -1, string Remark1 = "", string Remark2 = "", string Remark3 = "", string LCNO = "", string INVOICENO = ""
            , string DeclarationNumber = "", string NewRoFull = "", string ShipNoticeFull = "", string ChangeInvCode = ""
            , string DepositBatches = ""
            , string SoFull = "", string AdvOrderFull = "", double OffsetAmount = -1, double OffsetTaxAmount = -1, string TransportMethod = ""
            , string DispatchOrderFull = "", string CarNumber = "", string TrainNumber = "", string DeliveryUser = ""
            , string Courier = "", string SiteCommCalcMethod = "", string SiteCommRate = "", string TotalCommAmount = ""
            , string RoDetailData = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "update");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReceiveOrder(RoId, ViewCompanyId
                        , DocDate
                        , ReceiveDate, CustomerId, ProcessCode
                        , RevenueJournalCode, CostJournalCode, NoCredLimitControl, CashSales
                        , SourceType, PriceSourceTypeMain, SourceOrderId, SourceFull
                        , DepartmentId
                        , OriginalAmount, OriginalTaxAmount, TotalQuantity, PretaxAmount, TaxAmount
                        , TaxCurrencyRate
                        , SalesmenId, Factory, PaymentTerm, TradingTerms, Currency
                        , ExchangeRate, BondCode, RowCnt, Remark
                        , CustomerFullName, ContactPerson, CustomerAddressFirst, TelephoneNumber
                        , CustomerAddressSecond, FaxNumber
                        , TaxCode
                        , InvoiceType, TaxType, TaxRate, InvNumGenMethod, UiNo
                        , InvoiceDate, InvoiceTime, InvNumStart, InvNumEnd, ApplyYYMM, CustomsClearance
                        , CustomerFullNameOre, InvoiceAddressFirst, CustomerEgFullName, InvoiceAddressSecond
                        , MultipleInvoices, MultiTaxRate, AttachInvWithShip, InvoicesVoid
                        , VehicleTypeNumber
                        , InvDonationRecipient, VehicleIDshow, VehicleIDhide, CreditCard4No, InvRandomCode, ContactEmail
                        , StaffId
                        , CollectionSalesmenId, Remark1, Remark2, Remark3, LCNO, INVOICENO
                        , DeclarationNumber, NewRoFull, ShipNoticeFull, ChangeInvCode
                        , DepositBatches
                        , SoFull, AdvOrderFull, OffsetAmount, OffsetTaxAmount, TransportMethod
                        , DispatchOrderFull, CarNumber, TrainNumber, DeliveryUser
                        , Courier, SiteCommCalcMethod, SiteCommRate, TotalCommAmount
                        , RoDetailData
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

        #region //UpdateReceiveOrderToERP 銷貨單拋轉ERP
        [HttpPost]
        public void UpdateReceiveOrderToERP(int RoId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "import");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReceiveOrderToERP(RoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReceiveOrderReviseMES 銷貨單回歸MES重新編輯
        [HttpPost]
        public void UpdateReceiveOrderReviseMES(int RoId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "import");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReceiveOrderReviseMES(RoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoConfirm MES銷貨單確認單據
        [HttpPost]
        public void UpdateRoConfirm(int RoId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "confirm");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateRoConfirm(RoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoReconfirm MES銷貨單反確認單據
        [HttpPost]
        public void UpdateRoReconfirm(int RoId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "confirm");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateRoReconfirm(RoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReceiveOrderVoid 銷貨單作廢
        [HttpPost]
        public void UpdateReceiveOrderVoid(int RoId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "void");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReceiveOrderVoid(RoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReceiveOrderSynchronize 銷貨單資料手動同步
        [HttpPost]
        public void UpdateReceiveOrderSynchronize(string RoErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "", string CompanyNo = "")
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "sync");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReceiveOrderSynchronize(RoErpFullNo, SyncStartDate, SyncEndDate
                    , NormalSync, TranSync, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReturnReceiveOrderSynchronize 銷退單資料手動同步
        [HttpPost]
        public void UpdateReturnReceiveOrderSynchronize(string RtErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "", string CompanyNo = "")
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "sync");

                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReturnReceiveOrderSynchronize(RtErpFullNo, SyncStartDate, SyncEndDate
                    , NormalSync, TranSync, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReturnReceiveOrder 銷退單更新
        [HttpPost]
        public void UpdateReturnReceiveOrder(int RtId = -1, int ViewCompanyId = -1
            , string DocDate = ""
            , string ReturnDate = "", string RtErpPrefix = "", string RtErpNo = "", int CustomerId = -1, string ProcessCode = ""
            , string RevenueJournalCode = "", string CostJournalCode = ""
            , double OriginalAmount = -1
            , double OriginalTaxAmount = -1, double PretaxAmount = -1, double TaxAmount = -1, double TotalQuantity = -1
            , double TaxCurrencyRate = -1
            , int DepartmentId = -1
            , int SalesmenId = -1, int StaffId = -1, int CollectionSalesmenId = -1, string Factory = "", string PaymentTerm = "", string TradingTerms = ""
            , string Currency = "", double ExchangeRate = -1, string BondCode = "", int RowCnt = -1, string Remark = ""
            , string TaxCode = ""
            , string InvoiceType = "", string TaxType = "", double TaxRate = -1, string UiNo = "", string InvoiceDate = "", string InvNumStart = "", string InvNumEnd = ""
            , string ApplyYYMM = "", string CustomsClearance = ""
            , string CustomerFullName = ""
            , string CustomerEgFullName = "", string MultipleInvoices = "", string MultiTaxRate = "", string DebitNote = ""
            , string ContactPerson = ""
            , string ContactEmail = "", string CustomerAddressFirst = "", string CustomerAddressSecond = "", string Remark1 = "", string Remark2 = "", string Remark3 = ""
            , string RtDetailData = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "update");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReturnReceiveOrder(RtId,ViewCompanyId,
                            DocDate,
                            ReturnDate, RtErpPrefix, RtErpNo, CustomerId, ProcessCode,
                            RevenueJournalCode, CostJournalCode,
                            OriginalAmount,
                            OriginalTaxAmount, PretaxAmount, TaxAmount, TotalQuantity,
                            TaxCurrencyRate,
                            DepartmentId,
                            SalesmenId, StaffId, CollectionSalesmenId, Factory, PaymentTerm, TradingTerms,
                            Currency, ExchangeRate, BondCode, RowCnt, Remark,
                            TaxCode,
                            InvoiceType, TaxType, TaxRate, UiNo, InvoiceDate, InvNumStart, InvNumEnd,
                            ApplyYYMM, CustomsClearance,
                            CustomerFullName, CustomerEgFullName, MultipleInvoices, MultiTaxRate, DebitNote,
                            ContactPerson,
                            ContactEmail, CustomerAddressFirst, CustomerAddressSecond, Remark1, Remark2, Remark3,
                            RtDetailData
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

        #region //UpdateRtToERP 銷退單拋轉ERP
        [HttpPost]
        public void UpdateRtToERP(int RtId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "import");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateRtToERP(RtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRtReviseMES 銷退單回歸MES重新編輯
        [HttpPost]
        public void UpdateRtReviseMES(int RtId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "import");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateRtReviseMES(RtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRtVoid 銷退單作廢
        [HttpPost]
        public void UpdateRtVoid(int RtId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "void");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateRtVoid(RtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReturnRtConfirm 銷退單確認單據
        [HttpPost]
        public void UpdateReturnRtConfirm(int RtId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "confirm");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReturnRtConfirm(RtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReturnRtReconfirm 銷退單反確認單據
        [HttpPost]
        public void UpdateReturnRtReconfirm(int RtId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "confirm");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.UpdateReturnRtReconfirm(RtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteReceiveOrder 銷貨單刪除
        [HttpPost]
        public void DeleteReceiveOrder(int RoId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "delete");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.DeleteReceiveOrder(RoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRoDetail 銷貨單單身刪除
        [HttpPost]
        public void DeleteRoDetail(string RoDetailList = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "delete");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.DeleteRoDetail(RoDetailList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteReceiveOrderAllMES MES銷貨單全刪除
        [HttpPost]
        public void DeleteReceiveOrderAllMES(
            )
        {
            try
            {
                WebApiLoginCheck("ReceiveOrderManagment", "delete");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.DeleteReceiveOrderAllMES();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteReturnReceiveOrder 銷退單刪除
        [HttpPost]
        public void DeleteReturnReceiveOrder(int RtId = -1
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "delete");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.DeleteReturnReceiveOrder(RtId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteReturnRtDetail 銷貨單單身刪除
        [HttpPost]
        public void DeleteReturnRtDetail(string RtDetailList = ""
            )
        {
            try
            {
                WebApiLoginCheck("ReturnReceiveOrderManagment", "delete");


                #region //Request
                receiveOrderDA = new ReceiveOrderDA();
                dataRequest = receiveOrderDA.DeleteReturnRtDetail(RtDetailList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #endregion

    }
}