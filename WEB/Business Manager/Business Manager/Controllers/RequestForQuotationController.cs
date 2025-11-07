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
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Business_Manager.Controllers
{
    public class RequestForQuotationController : WebController
    {
        private RequestForQuotationDA requestForQuotationDA = new RequestForQuotationDA();      

        #region//View
        // GET: RequestForQuotation
        public ActionResult RfqManagment()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult RfqAssignManagment()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult RfqLineSolution()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QuotationManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QuotationWorkPlatform()
        {
            ViewLoginCheck();
            return View();
        }


        #endregion

        #region//Get
        #region //GetRequestForQuotation 取得客戶詢價資訊管理(RFQ)資訊
        [HttpPost]
        public void GetRequestForQuotation(int RfqId = -1, string RfqNo = "", int MemberId = -1, string MemberName = "", string AssemblyName = "", int ProductUseId = -1
            , int SalesId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "read,constrained-data");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetRequestForQuotation(RfqId, RfqNo, MemberId, MemberName, AssemblyName, ProductUseId
                    , SalesId, Status
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

        #region //GetRfqDetail 取得RFQ單身資訊
        [HttpPost]
        public void GetRfqDetail(int RfqId = -1, int RfqDetailId = -1, string DocType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetRfqDetail(RfqId, RfqDetailId, DocType
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

        #region //GetRfqDetailSales 取得RFQ單身資訊(負責業務)
        [HttpPost]
        public void GetRfqDetailSales(int RfqId = -1, int RfqDetailId = -1, string DocType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetRfqDetail(RfqId, RfqDetailId, DocType
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

        #region //GetSales 取得RFQ負責業務(Cmb用)
        [HttpPost]
        public void GetSales(int UserId = -1, string UserName = "", int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "read");

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

        #region //GetSalesQuery 取得RFQ負責業務(Query用)
        [HttpPost]
        public void GetSalesQuery(int UserId = -1, string UserName = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetSalesQuery(UserId, UserName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRfqLineSolution 取得RFQ報價方案資料
        [HttpPost]
        public void GetRfqLineSolution(int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqLineSolution", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetRfqLineSolution(RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetDfmQiSolution 取得RFQ報價資料
        [HttpPost]
        public void GetDfmQiSolution(int RfqLineSolutionId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqLineSolution", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetDfmQiSolution(RfqLineSolutionId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSalesDefault 取得RFQ系統預設客戶對應之業務人員
        [HttpPost]
        public void GetSalesDefault(int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetSalesDefault(RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQuotationDataModel 取得報價單樣板資料
        [HttpPost]
        public void GetQuotationDataModel(int RfqId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "read,constrained-data");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationDataModel(RfqId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQuotationDetail 取得報價單資料
        [HttpPost]
        public void GetQuotationDetail(int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "read,constrained-data");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationDetail( RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQuotationHistory 取得報價單歷史資料
        [HttpPost]
        public void GetQuotationHistory(int RfqDetailId = -1, int RfqProTypeId = -1, string TagList = "", string PreciseMode = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "read,constrained-data");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationHistory(RfqDetailId, RfqProTypeId, TagList, PreciseMode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetIncomeStatement 取得損益表
        [HttpPost]
        public void GetIncomeStatement(int RfqId = -1, string TagList = "", string MtlItemNo = "", string MtlItemName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-income");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetIncomeStatement(RfqId, TagList, MtlItemNo, MtlItemName,
                    OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQuotationAuthority 取得報價單相關權限
        [HttpPost]
        public void GetQuotationAuthority(string UserNo = "")
        {
            try
            {
                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationAuthority(UserNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQuotationBargainPrice 取得議價資料
        [HttpPost]
        public void GetQuotationBargainPrice(int RfqDetailId = -1)
        {
            try
            {
                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationBargainPrice(RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQuotationWork 取得報價單維護清單(工作平台列表)
        [HttpPost]
        public void GetQuotationWork(int RfqId = -1,string RfqNo = "", int CustomerId = -1, int RfqProTypeId = -1, string RoleMode = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "read");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationWork(RfqId, RfqNo, CustomerId, RfqProTypeId, RoleMode, Status
                    ,OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetAICycleTime AI自動計畫取得成形週期
        [HttpPost]
        public void GetAICycleTime(string Material = "",string Fi = "" ,string Thick = "")
        {
            try
            {
                if(Material.Length <=0) throw new Exception("【AI計算成形週期】材料不能為空");
                if (Fi.Length <=0) throw new Exception("【AI計算成形週期】外徑不能為空");
                if (Thick.Length <=0) throw new Exception("【AI計算成形週期】芯厚不能為空");

                string apiUrl = "http://192.168.20.97:6668/pmo_cycle"; // API網址上線後待切換
                var requestBody = new { material = Material, fi = Fi, thick = Thick };

                string MoldingCycle = "";
                string MaterialStatus = "";
                string MaterialErrMsg = "";
                // 使用 HttpClient 實例
                using (HttpClient client = new HttpClient())
                {
                    string json = JsonConvert.SerializeObject(requestBody);

                    // 建立 StringContent 作為請求內容
                    StringContent content = new StringContent(json, Encoding.UTF8, "application/json");

                    // 發送 POST 請求 (同步方式)
                    HttpResponseMessage response = client.PostAsync(apiUrl, content).Result; // 使用 .Result 來執行同步操作

                    // 確保成功回應
                    if (response.IsSuccessStatusCode)
                    {
                        string responseData = response.Content.ReadAsStringAsync().Result;

                        JObject parsedResponse = JObject.Parse(responseData);

                        string innerJsonString = parsedResponse["result"].ToString();

                        responseData = innerJsonString.Replace("'", "\"");

                        JObject parsedResponse1 = JObject.Parse(responseData);

                        string message = parsedResponse1["message"].ToString();
                        // 分割字符串為部分
                        var parts = message.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                        // 使用 LINQ 或簡化解析每個鍵值對
                        var result = parts
                            .Select(part =>
                            {
                                var keyValue = part.Split(':');
                                return new
                                {
                                    Key = keyValue[0].Trim(),
                                    Value = keyValue.Length > 1 ? keyValue[1].Replace("s", "").Trim() : null
                                };
                            })
                            .ToArray();
                        int num = 1;
                        foreach (var item in result)
                        {
                            if(num == 1)
                            {
                                MoldingCycle = item.Value;
                            }
                            if (num == 3)
                            {
                                MaterialStatus = (item.Value).Substring(0, 1) == "S" ? "Y" : "N";
                            }
                            num++;
                        }

                        //string[] parts = message.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                        //string test = parts[0];
                        //string[] parts1 = test.Split(new[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                        //string test2 = parts1[1].Replace("s", "").Trim();

                        if(MaterialStatus == "N")
                        {
                           MaterialErrMsg = $"當前材料{Material}為新材料,AI模型只能提供預測資料,請評估使用";
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = MoldingCycle,
                            msg = MaterialErrMsg
                        });
                        #endregion
                       
                    }
                    else
                    {
                        throw new Exception($"API 回應失敗，狀態碼: {response.StatusCode},請回報或是先採用人工輸入方式");
                    }
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

        #region //Add
        #region //AddRequestForQuotation RFQ單頭資料新增
        [HttpPost]
        public void AddRequestForQuotation(string RfqNo = "", string MemberName = "", string AssemblyName = "", int ProductUseId = -1
            , int MemberId = -1, string OrganizaitonType = "", int CustomerId = -1, string CustomerName = "", int SupplierId = -1, string SupplierName = ""
            , string ContactPerson = "", string ContactPhone = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "add");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.AddRequestForQuotation(RfqNo, MemberName, AssemblyName, ProductUseId
                    , MemberId, OrganizaitonType, CustomerId, CustomerName, SupplierId, SupplierName
                    , ContactPerson, ContactPhone);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRfqDetail RFQ單身資料新增
        [HttpPost]
        public void AddRfqDetail(int RfqId = -1, int CompanyId = -1, int RfqProTypeId = -1, string RfqSequence = "", string MtlName = ""
            , int CustProdDigram = -1, string PlannedOpeningDate = "", int PrototypeQty = -1, string ProtoSchedule = "", string MassProductionDemand = ""
            , string KickOffType = "", string PlasticName = "", string OutsideDiameter = "", string ProdLifeCycleStart = "", string ProdLifeCycleEnd = ""
            , int LifeCycleQty = -1, string DemandDate = "", string CoatingFlag = "", string Currency = "", int MonthlyQty = -1, int SampleQty = -1, int UomId = -1, string PortOfDelivery = "", int SalesId = -1
            , int AdditionalFile = -1, int QuotationFile = -1, string Description = "", string QuotationRemark = "", string ConfirmVPTime = "", string ConfirmSalesTime = ""
            , string ConfirmRdTime = "", string ConfirmCustTime = "", string Edition = "", string DocType = "", string Status = ""
            , string DetailStatus = "", string PackagingDetails = "", string SolutionInfos = "", int BaseCavities = -1, int InsertCavities = -1, string CoreThickness = "", string CommonMode = "", string DesignChange = "")
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "add");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.AddRfqDetail(RfqId, CompanyId, RfqProTypeId, RfqSequence, MtlName
                    , CustProdDigram, PlannedOpeningDate, PrototypeQty, ProtoSchedule, MassProductionDemand
                    , KickOffType, PlasticName, OutsideDiameter, ProdLifeCycleStart, ProdLifeCycleEnd
                    , LifeCycleQty, DemandDate, CoatingFlag, Currency, MonthlyQty, SampleQty, UomId, PortOfDelivery, SalesId
                    , AdditionalFile, QuotationFile, Description, QuotationRemark, ConfirmVPTime, ConfirmSalesTime
                    , ConfirmRdTime, ConfirmCustTime, Edition, DocType, Status
                    , DetailStatus, PackagingDetails, SolutionInfos, BaseCavities, InsertCavities, CoreThickness, CommonMode, DesignChange);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRfqDetailDocType 設變流程(複製RFQ資料功能)
        [HttpPost]
        public void AddRfqDetailDocType(int RfqId = -1, int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "add");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.AddRfqDetailDocType(RfqId, RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQuotationBargain 新增報價議價
        [HttpPost]
        public void AddQuotationBargain(int RfqDetailId = -1, double FinalPrice = -1, double GrossMargin = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "add");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.AddQuotationBargain(RfqDetailId, FinalPrice, GrossMargin);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateRequestForQuotation RFQ單頭資料更新
        [HttpPost]
        public void UpdateRequestForQuotation(int RfqId = -1, string MemberName = "", string AssemblyName = "", int ProductUseId = -1
            , string OrganizaitonType = "", int CustomerId = -1, string CustomerName = "", int SupplierId = -1, string SupplierName = ""
            , string ContactPerson = "", string ContactPhone = "")
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRequestForQuotation(RfqId, MemberName, AssemblyName, ProductUseId
                    , OrganizaitonType, CustomerId, CustomerName, SupplierId, SupplierName
                    , ContactPerson, ContactPhone);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqSales 主頁批次指派業務更新
        [HttpPost]
        public void UpdateRfqSales(int RfqId = -1, int CompanyId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRfqSales(RfqId, CompanyId, UserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqDetail RFQ單身明細更新
        [HttpPost]
        public void UpdateRfqDetail(int RfqId = -1, int RfqProTypeId = -1, int RfqDetailId = -1
            , string MtlName = "", string PlannedOpeningDate = "", int PrototypeQty = -1, string ProtoSchedule = "", string MassProductionDemand = ""
            , string KickOffType = "", string PlasticName = "", string OutsideDiameter = "", string ProdLifeCycleStart = "", string ProdLifeCycleEnd = ""
            , int LifeCycleQty = -1, string DemandDate = "", string CoatingFlag = "", string Currency = "", int MonthlyQty = -1,int SampleQty = -1, int UomId = -1, int CustProdDigram = -1, int AdditionalFile = -1
            , int QuotationFile = -1, string Description = "", string PackagingDetails = "", string SolutionInfos = "", int BaseCavities = -1, int InsertCavities = -1, string CoreThickness = "", string CommonMode = "")
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRfqDetail(RfqId, RfqProTypeId, RfqDetailId
                    , MtlName, PlannedOpeningDate, PrototypeQty, ProtoSchedule, MassProductionDemand
                    , KickOffType, PlasticName, OutsideDiameter, ProdLifeCycleStart, ProdLifeCycleEnd
                    , LifeCycleQty, DemandDate, CoatingFlag, Currency, MonthlyQty, SampleQty, UomId, CustProdDigram, AdditionalFile
                    , QuotationFile, Description, PackagingDetails, SolutionInfos, BaseCavities, InsertCavities, CoreThickness, CommonMode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqDetailSales 單身單筆指派業務更新
        [HttpPost]
        public void UpdateRfqDetailSales(int RfqDetailId = -1, int CompanyId = -1, int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRfqDetailSales(RfqDetailId, CompanyId, UserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqConfirm 主頁批次確認狀態更新
        [HttpPost]
        public void UpdateRfqConfirm(int RfqId = -1, int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRfqConfirm(RfqId, RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqDetailConfirm 單身單筆確認狀態更新
        [HttpPost]
        public void UpdateRfqDetailConfirm(int RfqId = -1, int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRfqDetailConfirm(RfqId, RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDfmQiSolution RFQ單身報價資料更新(單筆儲存)
        [HttpPost]
        public void UpdateDfmQiSolution(int DfmQiId = -1, int RfqLineSolutionId = -1
           , double DiscountAmount = 0.0, double GrossProfitMargin = 0.0, double AfterProfitMargin = 0.0, double QuotationAmount = 0.0)
        {
            try
            {
                WebApiLoginCheck("RfqLineSolution", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateDfmQiSolution(DfmQiId, RfqLineSolutionId
                    , DiscountAmount, GrossProfitMargin, AfterProfitMargin, QuotationAmount);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLotDfmQiSolution RFQ單身報價資料更新(多筆儲存)
        [HttpPost]
        public void UpdateLotDfmQiSolution(string Quotations = "")
        {
            try
            {
                WebApiLoginCheck("RfqLineSolution", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateLotDfmQiSolution(Quotations);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDfmQiSolutionConfirm RFQ報價確認狀態更新
        [HttpPost]
        public void UpdateDfmQiSolutionConfirm(int RfqId = -1, int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqLineSolution", "update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateDfmQiSolutionConfirm(RfqId, RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRfqQuotationTag RFQ報價單標籤更新
        [HttpPost]
        public void UpdateRfqQuotationTag(int RfqDetailId = -1, string TagList = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateRfqQuotationTag(RfqDetailId, TagList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationRemark RFQ報價單備註更新
        [HttpPost]
        public void UpdateQuotationRemark(int RfqDetailId = -1, string QuotationRemark = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateQuotationRemark(RfqDetailId, QuotationRemark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationDetail 報價單項目資料更新
        [HttpPost]
        public void UpdateQuotationDetail(int QiId = -1, string Data = "", string AiData = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-update");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateQuotationDetail(QiId, Data, AiData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationDetailConfirm 報價單項目資料確認/反確認
        [HttpPost]
        public void UpdateQuotationDetailConfirm(int QiId = -1, string Confirm = "", bool ClassLast = false, string Data = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-confirm");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateQuotationDetailConfirm(QiId, Confirm, ClassLast, Data);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationFinalConfirm 報價單最終確認
        [HttpPost]
        public void UpdateQuotationFinalConfirm(int QiId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-confirm");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateQuotationFinalConfirm(QiId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationBargainConfirm 報價單議價確認
        [HttpPost]
        public void UpdateQuotationBargainConfirm(int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-review");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateQuotationBargainConfirm(RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQuotationBargainReview 報價單議價確認
        [HttpPost]
        public void UpdateQuotationBargainReview(int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-review");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.UpdateQuotationBargainReview(RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteRequestForQuotation -- 刪除RFQ單頭資料
        [HttpPost]
        public void DeleteRequestForQuotation(int RfqId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "delete");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.DeleteRequestForQuotation(RfqId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRfqDetail 刪除RFQ單身資料
        [HttpPost]
        public void DeleteRfqDetail(int RfqDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("RfqManagment", "delete");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.DeleteRfqDetail(RfqDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //API
        #region //NotifyQuotationQuestion 報價項目異議通知
        [HttpPost]
        public void NotifyQuotationQuestion(int QiId = -1, string Remark ="")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-notify");

                #region //Request
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.NotifyQuotationQuestion(QiId, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region//Word
        #region//QuotationDocDownload報價單

        public void QuotationDocDownload(int RfqDetailId = -1, string Remark = "") {
            try {
                
                requestForQuotationDA = new RequestForQuotationDA();
                dataRequest = requestForQuotationDA.GetQuotationHead(RfqDetailId);
                jsonResponse = JObject.Parse(dataRequest);

                double ExchangeEate = 0.0;

                if (jsonResponse["status"].ToString() == "success") {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("供應商資料未維護完整!");

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
                            wordFileName = "Quotation";
                            filePath = "~/WordTemplate/EIP/Jmo_Quotation.docx";
                            break;
                        case 4: //晶彩
                            wordFileName = "Quotation";
                            filePath = "~/WordTemplate/EIP/Jc_Quotation.docx";
                            break;
                        default:
                            throw new SystemException("目前公司別未開放功能!");
                    }
                    #endregion

                    var dfmqisolution = result["data"];
                    using (DocX doc = DocX.Load(Server.MapPath(filePath))) {
                        #region//單頭
                        doc.ReplaceText("[CustomerName]", dfmqisolution[0]["CustomerName"].ToString());
                        doc.ReplaceText("[Contact]", dfmqisolution[0]["Contact"].ToString());
                        doc.ReplaceText("[TelNoFirst]", dfmqisolution[0]["TelNoFirst"].ToString());
                        doc.ReplaceText("[GuiNumber]", dfmqisolution[0]["GuiNumber"].ToString());
                        doc.ReplaceText("[RegisterAddressFirst]", dfmqisolution[0]["RegisterAddressFirst"].ToString());
                        doc.ReplaceText("[FaxNo]", dfmqisolution[0]["FaxNo"].ToString());
                        doc.ReplaceText("[RFQ011]", "");
                        doc.ReplaceText("[RFQ012]", "");
                        doc.ReplaceText("[RFQ013]", "");
                        doc.ReplaceText("[RFQ014]", "");
                        doc.ReplaceText("[RFQ015]", "");
                        #endregion

                        #region//單身
                        requestForQuotationDA = new RequestForQuotationDA();
                        dataRequest = requestForQuotationDA.GetRfqLineSolution(RfqDetailId);
                        jsonResponse = JObject.Parse(dataRequest);
                        if (jsonResponse["status"].ToString() == "success")
                        {
                            //取得方案
                            for (int i = 0; i < jsonResponse["data"].Count(); i++)
                            {
                                doc.ReplaceText("[SolutionQty" + i + "]", jsonResponse["data"][i]["SolutionQty"].ToString());
                                doc.ReplaceText("[PeriodicDemandType" + i + "]", "pcs/月(成套)");
                                doc.ReplaceText("[Currency" + i + "]", jsonResponse["data"][i]["Currency"].ToString());
                                doc.ReplaceText("[RFQ016]", jsonResponse["data"][i]["Currency"].ToString());

                                //取得ERP匯率
                                var ExchangeRateRequest = requestForQuotationDA.GetExchangeRate(jsonResponse["data"][i]["Currency"].ToString());
                                var ExchangeRateJsonResponse = JObject.Parse(ExchangeRateRequest);
                                if (ExchangeRateJsonResponse["status"].ToString() != "success") throw new SystemException("查無匯率資訊!");
                                for (int a = 0; a < ExchangeRateJsonResponse["data"].Count(); a++) {
                                    ExchangeEate = Convert.ToDouble(ExchangeRateJsonResponse["data"][a]["CustomsSellingRate"].ToString());//暫時抓【報關賣出匯率】
                                }                                    
                            }
                            for (int s= jsonResponse["data"].Count();s<5;s++) {
                                doc.ReplaceText("[SolutionQty" + s + "]", "");
                                doc.ReplaceText("[PeriodicDemandType" + s + "]", "");
                                doc.ReplaceText("[Currency" + s + "]", "");
                            }
                        }

                        requestForQuotationDA = new RequestForQuotationDA();
                        dataRequest = requestForQuotationDA.GetQuotationLine(RfqDetailId);
                        jsonResponse = JObject.Parse(dataRequest);
                        if (jsonResponse["status"].ToString() == "success")
                        {
                            for (int i = 0; i < jsonResponse["data"].Count(); i++)
                            {
                                int z = 0;//品號數                             
                                for (z = jsonResponse["data"].Count(); z <= 18; z++)
                                {
                                    doc.ReplaceText("[MtlItemName" + z + "]", "");
                                    for (int y = 0; y <= 4; y++)
                                    {
                                        doc.ReplaceText("[QuotationAmount-" + y + "," + z + "]", "");
                                    }
                                }
                                doc.ReplaceText("[MtlItemName" + i + "]", jsonResponse["data"][i]["DfmQuotationName"].ToString());
                                string a = jsonResponse["data"][i]["Detail"].ToString();
                                JObject Detail = JObject.Parse(a);
                                int w = Detail["data"].Count();
                                for (int j = 0; j < Detail["data"].Count(); j++)
                                {
                                    //金額乘上匯率
                                    double QuotationAmount = Math.Round(Convert.ToDouble(Detail["data"][j]["QuotationAmount"].ToString())/ExchangeEate,3);
                                    doc.ReplaceText("[QuotationAmount-" + j + "," + i + "]", QuotationAmount.ToString());
                                }

                                for (int p= Detail["data"].Count(); p <= 4; p++)
                                {
                                    for (int t = 0; t <= 18; t++)
                                    {
                                        doc.ReplaceText("[QuotationAmount-" + p + "," +t+ "]", "");
                                    }
                                }
                            }
                        }
                        #endregion

                        #region//備註
                        doc.ReplaceText("[Remark]", Remark);
                        #endregion

                        wordFileName = dfmqisolution[0]["QuotationFileName"].ToString();

                        using (MemoryStream output = new MemoryStream())
                        {
                            doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                            doc.SaveAs(output, password);

                            wordFile = output.ToArray();
                        }
                    }

                    fileGuid = Guid.NewGuid().ToString();
                    Session[fileGuid] = wordFile;

                    #region//word存入file並更新SCM.RfqDetail
                    var fileBinary = wordFile;
                    string FileName = wordFileName;
                    string FileExtension =".docx";
                    int FileSize = Convert.ToInt32(wordFile.Length);
                    string Source = "/RequestForQuotation/RfqLineSolution";
                    string ClientIP = "::1";

                    #region //寫入資料庫
                    dataRequest = requestForQuotationDA.AddQuotationFile(fileBinary, FileName, FileExtension, FileSize, ClientIP, Source);
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    int FileId = int.Parse(jsonResponse["result"][0]["FileId"].ToString());
                    #endregion

                    #region//Update RfqDetial
                    dataRequest = requestForQuotationDA.UpdateRfqDetailQuotation(FileId,RfqDetailId);
                    jsonResponse = JObject.Parse(dataRequest);
                    #endregion

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

        #region //GetQuotationPdf 報價單輸出
        public void GetQuotationPdf(int RfqId = -1, int RfqDetailId = -1, string RfqIdList = "", string TaxIncluded = "", string ExchangeStatus = "")
        {
            try
            {
                WebApiLoginCheck("RfqAssignManagment", "quote-print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - 報價單PDF資料
                dataRequest = requestForQuotationDA.GetQuotationPdf(RfqId,RfqDetailId, RfqIdList, ExchangeStatus);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                string QuotationNo = "";
                string htmlText = "";

                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("此製令無法進行列印作業!可能因圖層未設定");

                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/QuotationPdf.html"));
                    string itemDetail = "";
                    int row = 1;
                    double TotalAmount = 0;
                    foreach (var itme in result)
                    {
                        string MtlName = itme["MtlName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string InsertCavities = itme["InsertCavities"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string UomNo = itme["UomNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string MonthlyQty = itme["MonthlyQty"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string UnitPrice = Convert.ToDouble(itme["UnitPrice"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ")).ToString("N4");
                        string SetupCost = itme["SetupCost"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string Amount = "";
                        string QuotationRemark = itme["QuotationRemark"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                        string ExchangeRate = itme["ExchangeRate"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                        if (row == 1)
                        {
                            QuotationNo = itme["QuotationNo"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string CustomerName = itme["CustomerName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string ContactPhone = itme["ContactPhone"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string ContactPerson = itme["ContactPerson"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string QuotationDate = itme["QuotationDate"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string Currency = itme["Currency"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string Taxation = itme["Taxation"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string TaxRate = itme["TaxRate"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string PaymentTermName = itme["PaymentTermName"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");
                            string CreateUser = itme["CreateUser"].ToString().Replace("<", " &#60; ").Replace("&", " &#38; ");

                            switch (Taxation)
                            {
                                case "1":
                                    Taxation = "應稅內含";
                                    UnitPrice = (Convert.ToDouble(UnitPrice) * (1 + Convert.ToDouble(TaxRate))).ToString("N4");
                                    break;
                                case "2":
                                    Taxation = "應稅外加";
                                    UnitPrice = (Convert.ToDouble(UnitPrice) + (Convert.ToDouble(UnitPrice) * Convert.ToDouble(TaxRate))).ToString("N4");
                                    break;
                                case "3":
                                    Taxation = "零稅率";
                                    break;
                                case "4":
                                    Taxation = "免稅";
                                    break;
                                case "9":
                                    Taxation = "不計稅";
                                    break;
                            }

                            htmlText = htmlText.Replace("[Web]", Server.MapPath("~/PdfTemplate/MES/Quotation/DGJMO.jpg"));
                            htmlText = htmlText.Replace("[CustomerName]", CustomerName);
                            htmlText = htmlText.Replace("[ContactPhone]", ContactPhone);
                            htmlText = htmlText.Replace("[ContactPerson]", ContactPerson);
                            htmlText = htmlText.Replace("[QuotationNo]", QuotationNo);
                            htmlText = htmlText.Replace("[QuotationDate]", QuotationDate);
                            htmlText = htmlText.Replace("[Currency]", Currency);
                            htmlText = htmlText.Replace("[Taxation]", Taxation);
                            htmlText = htmlText.Replace("[CreateUser]", CreateUser);
                            htmlText = htmlText.Replace("[PaymentTermName]", PaymentTermName);
                            htmlText = htmlText.Replace("[ExchangeRate]", ExchangeRate);

                        }

                        UnitPrice = ExchangeStatus == "Y" ? (Convert.ToDouble(UnitPrice) / Convert.ToDouble(ExchangeRate)).ToString("N4") : UnitPrice;

                        Amount = (Convert.ToDouble(UnitPrice) * (Convert.ToDouble(MonthlyQty)) + (Convert.ToDouble(SetupCost))).ToString("N4");

                        string detail = @"<tr style=""height: 30px; "">
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[Row]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[MtlName]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[InsertCavities]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[UomNo]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[MonthlyQty]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[UnitPrice]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[SetupCost]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[Amount]</td>
                                            <td colspan=""1"" style=""font-size:12px;text-align: center;"">[QuotationRemark]</td>
                                        </tr>";
                        detail = detail.Replace("[Row]", row.ToString());
                        detail = detail.Replace("[MtlName]", MtlName);
                        detail = detail.Replace("[InsertCavities]", InsertCavities);
                        detail = detail.Replace("[UomNo]", UomNo);
                        detail = detail.Replace("[MonthlyQty]", MonthlyQty);
                        detail = detail.Replace("[UnitPrice]", UnitPrice);
                        detail = detail.Replace("[SetupCost]", SetupCost);
                        detail = detail.Replace("[Amount]", Amount);
                        detail = detail.Replace("[QuotationRemark]", QuotationRemark);
                        itemDetail += detail;
                        TotalAmount += Convert.ToDouble(Amount);
                        row++;
                    }

                    #region //html
                    htmlText = htmlText.Replace("[itemDetail]", itemDetail);
                    htmlText = htmlText.Replace("[TotalAmount]", TotalAmount.ToString("N4"));
                    string htmlTemplate = htmlText;
                    #endregion
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
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

                            try
                            {
                                XMLWorkerHelper.GetInstance().ParseXHtml(
                                    pdfWriter,
                                    document,
                                    input,
                                    null,
                                    Encoding.UTF8,
                                    new UnicodeFontFactory() // 確保正確加載字型，特別是中文
                                );
                            }
                            catch (Exception ex)
                            {
                                // 錯誤處理
                                Console.WriteLine("Error while parsing HTML to PDF: " + ex.Message);
                            }

                            //XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, document, input, null, Encoding.UTF8, new UnicodeFontFactory()); //使用XMLWorkerHelper把Html parse到PDF檔裡

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
                    fileName = "【報價單】" + QuotationNo,
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

    }
}