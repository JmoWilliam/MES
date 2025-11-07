using ClosedXML.Excel;
using Helpers;
using MESDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls.Expressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Business_Manager.Controllers
{
    public class ScmInventoryController : WebController
    {
        private ScmInventoryDA scmInventoryDA = new ScmInventoryDA();
        private SaleOrderDA saleOrderDA = new SaleOrderDA();

        #region //View
        public ActionResult InventoryTransaction()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TempShippingNote()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TempShippingReturnNote()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult InventoryQuery()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult InventoryAgingReport()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetInventoryTransaction 取得庫存異動資料
        [HttpPost]
        public void GetInventoryTransaction(int ItId = -1, string ItErpPrefix = "", string ItErpNo = "", string SearchKey = ""
            , string ConfirmStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetInventoryTransaction(ItId, ItErpPrefix, ItErpNo, SearchKey
                    , ConfirmStatus, StartDate, EndDate
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

        #region //GetTempShippingNote 取得暫出單資料
        [HttpPost]
        public void GetTempShippingNote(int TsnId = -1, string TsnErpPrefix = "", string TsnErpNo = "", string SearchKey = ""
            , int ObjectCustomer = -1, string ConfirmStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetTempShippingNote(TsnId, TsnErpPrefix, TsnErpNo, SearchKey
                    , ObjectCustomer, ConfirmStatus, StartDate, EndDate
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

        #region //GetDeliveryToTsn 取得出貨轉暫出資料
        [HttpPost]
        public void GetDeliveryToTsn(string SearchKey = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "delivery-to-tsn");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetDeliveryToTsn(SearchKey, StartDate, EndDate
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

        #region //GetTempShippingReturnNote 取得暫出歸還單資料
        [HttpPost]
        public void GetTempShippingReturnNote(int TsrnId = -1, string TsrnErpPrefix = "", string TsrnErpNo = "", string SearchKey = ""
            , string ConfirmStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TempShippingReturnNote", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetTempShippingReturnNote(TsrnId, TsrnErpPrefix, TsrnErpNo, SearchKey
                    , ConfirmStatus, StartDate, EndDate
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

        #region //GetInventoryQuery 庫存整合查詢
        [HttpPost]
        public void GetInventoryQuery(string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("InventoryQuery", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetInventoryQuery(SearchKey
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

        #region //GetPriceSwitch 顯示金額切換
        [HttpPost]
        public void GetPriceSwitch()
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "price-switch");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = ""
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

        #region //GetTsnDocDate 取得暫出單單據日期資料
        [HttpPost]
        public void GetTsnDocDate(int TsnId = -1, string ConfirmStatus = "")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetTsnDocDate(TsnId, ConfirmStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSoSalesInfo 取得訂單及暫出單人員資料
        [HttpPost]
        public void GetSoSalesInfo(string DoDetailId = "")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetSoSalesInfo(DoDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region /GetInventoryAgingReport 取得存貨庫齡分析表
        [HttpPost]
        public void GetInventoryAgingReport(string MtlItemNo = "", string AgingDate = "", string ChangeType = "", string DataDate = ""
            , int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("RptMoInformation", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetInventoryAgingReport(MtlItemNo, AgingDate, ChangeType, DataDate
                    , PageIndex, PageSize);
                #endregion

                #region //Response
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddDeliveryToTsn 出貨轉暫出資料新增
        [HttpPost]
        public void AddDeliveryToTsn(string TsnErpPrefix = "", int SalesmenId = -1, int DepartmentId = -1, string DoDetails = "",string Remark="")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "delivery-to-tsn");

                #region//暫出單檢核【信用額度】

                #endregion


                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.AddDeliveryToTsn(TsnErpPrefix, SalesmenId, DepartmentId, DoDetails,Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeliveryToTsn1 出貨轉暫出資料新增
        [HttpPost]
        public void AddDeliveryToTsn1(string TsnErpPrefix = "", int SalesmenId = -1, int DepartmentId = -1, string DoDetails = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "delivery-to-tsn");

                #region//暫出單檢核【信用額度】

                #endregion


                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.AddDeliveryToTsn1(TsnErpPrefix, SalesmenId, DepartmentId, DoDetails, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateInventoryTransactionManualSynchronize 庫存異動單資料手動同步
        [HttpPost]
        public void UpdateInventoryTransactionManualSynchronize(string ItErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "")
        {
            try
            {
                WebApiLoginCheck("InventoryTransaction", "sync");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.UpdateInventoryTransactionManualSynchronize(ItErpFullNo, SyncStartDate, SyncEndDate
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

        #region //UpdateTempShippingNoteManualSynchronize 暫出單資料手動同步
        [HttpPost]
        public void UpdateTempShippingNoteManualSynchronize(string TsnErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "sync");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.UpdateTempShippingNoteManualSynchronize(TsnErpFullNo, SyncStartDate, SyncEndDate
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

        #region //UpdateTempShippingReturnNoteManualSynchronize 暫出歸還單資料手動同步
        [HttpPost]
        public void UpdateTempShippingReturnNoteManualSynchronize(string TsrnErpFullNo = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "sync");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.UpdateTempShippingReturnNoteManualSynchronize(TsrnErpFullNo, SyncStartDate, SyncEndDate
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

        #region //UpdateTsnDocDate 暫出單單據日期資料更新
        [HttpPost]
        public void UpdateTsnDocDate(int TsnId = -1, string DocDate = "")
        {
            try
            {
                WebApiLoginCheck("TempShippingNote", "doc-date");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.UpdateTsnDocDate(TsnId, DocDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #endregion

        #region //Api
        #region //UpdateInventoryTransactionSynchronize 庫存異動資料同步
        [HttpPost]
        [Route("api/ERP/InventoryTransactionSynchronize")]
        public void UpdateInventoryTransactionSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateInventoryTransactionSynchronize");
                #endregion

                #region //Request
                dataRequest = scmInventoryDA.UpdateInventoryTransactionSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTempShippingNoteSynchronize 暫出單資料同步
        [HttpPost]
        [Route("api/ERP/TempShippingNoteSynchronize")]
        public void UpdateTempShippingNoteSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateTempShippingNoteSynchronize");
                #endregion

                #region //Request
                dataRequest = scmInventoryDA.UpdateTempShippingNoteSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTempShippingReturnNoteSynchronize 暫出歸還單資料同步
        [HttpPost]
        [Route("api/ERP/TempShippingReturnNoteSynchronize")]
        public void UpdateTempShippingReturnNoteSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateTempShippingReturnNoteSynchronize");
                #endregion

                #region //Request
                dataRequest = scmInventoryDA.UpdateTempShippingReturnNoteSynchronize(Company, UpdateDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //Word 
        #region //TsnDocDownload 暫出單(出貨單)資料單據下載
        public void TsnDocDownload(int TsnId = -1, string NonPrice = "", string LotSetting = "", string ItemMerge="")
        {
            try
            {
                if (!Regex.IsMatch(NonPrice, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【單據金額選項】錯誤!");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetTempShippingNoteDoc(TsnId, LotSetting);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("單據資料不完整!");

                    bool showPrice = true;
                    switch (NonPrice)
                    {
                        case "N":
                            showPrice = false;
                            break;
                        case "Y":
                            showPrice = true;
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
                            wordFileName = "出貨單-{0}-{1}";
                            filePath = "~/WordTemplate/SCM/暫出入單憑證-中揚.docx";
                            secondFilePath = "~/WordTemplate/SCM/暫出入單憑證-中揚-換頁.docx";
                            break;
                        case 4: //晶彩
                            wordFileName = "出貨單-{0}-{1}";
                            filePath = "~/WordTemplate/SCM/R-SD01表09 出貨單-晶彩.docx";
                            secondFilePath = "~/WordTemplate/SCM/R-SD01表09 出貨單-晶彩 換頁.docx";
                            break;
                        default:
                            throw new SystemException("目前公司別未開放功能!");
                    }
                    #endregion

                    var invtf = result["data"];
                    var invtg = result["dataDetail"];

                    #region //產生Doc
                    
                    int CompanydocRowH = 0;
                    if (CurrentCompany == 2)
                    {
                        CompanydocRowH = 12;
                    }
                    else
                    {
                        CompanydocRowH = 9;
                    }
                    totalPage = invtg.Count() / CompanydocRowH + (invtg.Count() % CompanydocRowH > 0 ? 1 : 0);

                    if (totalPage == 1)
                    {
                        #region //單頁
                        using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                        {
                            wordFileName = string.Format(wordFileName, invtf[0]["TF001"].ToString(), invtf[0]["TF002"].ToString());

                            #region //單頭
                            doc.ReplaceText("[MQ002]", invtf[0]["MQ002"].ToString());               //單據名稱
                            doc.ReplaceText("[TF001Doc]", invtf[0]["TF001Doc"].ToString());         //异动单别
                            doc.ReplaceText("[TF002]", invtf[0]["TF002"].ToString());               //异动单号
                            doc.ReplaceText("[TF024Date]", invtf[0]["TF024Date"].ToString());       //单据日期
                            doc.ReplaceText("[TF004]", invtf[0]["TF004"].ToString());               //异动对象
                            doc.ReplaceText("[TF015]", invtf[0]["TF015"].ToString());               //对象全名
                            doc.ReplaceText("[TF016]", invtf[0]["TF016"].ToString());               //转入地址
                            doc.ReplaceText("[TF018]", invtf[0]["TF018"].ToString());               //其它备注
                            doc.ReplaceText("[TF007Dep]", invtf[0]["TF007Dep"].ToString());         //部门
                            doc.ReplaceText("[TF008Staff]", invtf[0]["TF008Staff"].ToString());     //员工
                            doc.ReplaceText("[TF009Fac]", invtf[0]["TF009Fac"].ToString());         //厂别
                            doc.ReplaceText("[TF005]", invtf[0]["TF005"].ToString());               //对象代码
                            doc.ReplaceText("[TF011]", invtf[0]["TF011"].ToString());               //币别

                            if (CurrentCompany == 2)
                            {
                                switch (invtf[0]["TF010"].ToString()) {
                                    case "外加":
                                        doc.ReplaceText("[TF010]", "應稅外加");               //课税别
                                        break;
                                    case "內含":
                                        doc.ReplaceText("[TF010]", "應稅內含");               //课税别
                                        break;
                                    default:
                                        doc.ReplaceText("[TF010]", invtf[0]["TF010"].ToString());
                                        break;
                                }
                            }
                            else {
                                doc.ReplaceText("[TF010]", invtf[0]["TF010"].ToString());               //课税别
                            }
                           
                            doc.ReplaceText("[TF013]", invtf[0]["TF013"].ToString());               //件数
                            doc.ReplaceText("[TF014]", invtf[0]["TF014"].ToString());               //备注
                            doc.ReplaceText("[TF012Rate]", invtf[0]["TF012Rate"].ToString());       //汇率
                            doc.ReplaceText("[TF026]", invtf[0]["TF026"].ToString());               //营业税率
                            doc.ReplaceText("[TF020]", invtf[0]["TF020"].ToString());               //确认

                            DateTime currentDate = DateTime.Now;
                            string formattedDate = currentDate.ToString("yyyy/MM/dd");
                            doc.ReplaceText("[CreateDate]", formattedDate);     //製表日期
                            

                            doc.ReplaceText("[cp]", "1");
                            doc.ReplaceText("[tp]", "1");

                            if (showPrice)
                            {
                                doc.ReplaceText("[TF023Amount]", invtf[0]["TF023Amount"].ToString());   //暂出(入)金额
                                doc.ReplaceText("[TF027Tax]", invtf[0]["TF027Tax"].ToString());         //税额
                                doc.ReplaceText("[TotalAmount]", invtf[0]["TotalAmount"].ToString());   //金额合计
                            }
                            else
                            {
                                doc.ReplaceText("[TF023Amount]", "");
                                doc.ReplaceText("[TF027Tax]", "");
                                doc.ReplaceText("[TotalAmount]", "");
                            }

                            doc.ReplaceText("[TF022Qua]", invtf[0]["TF022Qua"].ToString());
                            
                            #endregion

                            #region //單身
                            if (invtg.Count() > 0)
                            {
                                for (int i = 0; i < invtg.Count(); i++)
                                {
                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invtg[i]["Seq"].ToString());                   //序号
                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invtg[i]["MtlItemNo"].ToString());       //品号

                                    string tempMtlItemName = invtg[i]["MtlItemName"].ToString();                                        //品名

                                    if (CurrentCompany==2) {
                                        if (tempMtlItemName.Length > 26)
                                        {
                                            tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 25) + "...";
                                        }
                                    }
                                    else {
                                        if (tempMtlItemName.Length > 20)
                                        {
                                            tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 20) + "...";
                                        }
                                    }


                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invtg[i]["MtlItemSpec"].ToString());   //规格

                                    doc.ReplaceText(string.Format("[OutStock{0:00}]", i + 1), invtg[i]["OutStock"].ToString());         //转出库别
                                    doc.ReplaceText(string.Format("[InStock{0:00}]", i + 1), invtg[i]["InStock"].ToString());           //转入库别
                                    doc.ReplaceText(string.Format("[ReturnDate{0:00}]", i + 1), invtg[i]["ReturnDate"].ToString());     //预计归还日
                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invtg[i]["Qty"].ToString());                   //数量
                                    doc.ReplaceText(string.Format("[PsQty{0:00}]", i + 1), invtg[i]["PsQty"].ToString());               //转进销量
                                    doc.ReplaceText(string.Format("[GpQty{0:00}]", i + 1), invtg[i]["GpQty"].ToString());               //贈/備品量
                                    doc.ReplaceText(string.Format("[RQty{0:00}]", i + 1), invtg[i]["RQty"].ToString());                 //归还量
                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invtg[i]["Uom"].ToString());                   //单位
                                    doc.ReplaceText(string.Format("[Muom{0:00}]", i + 1), invtg[i]["Muom"].ToString());                 //小单位

                                    if (CurrentCompany == 2){
                                        switch (invtg[i]["Closeout"].ToString()) {
                                            case "N":
                                                doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "未結案");         //结案码
                                                break;
                                            case "Y":
                                                doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "已結案");         //结案码
                                                break;
                                        }
                                    }
                                    else {
                                        doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), invtg[i]["Closeout"].ToString());         //结案码

                                    }

                                    if (showPrice)
                                    {
                                        doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), invtg[i]["UnPri"].ToString());           //单价
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), Convert.ToDouble(invtg[i]["Amount"]).ToString("n0"));   //金额
                                    }
                                    else
                                    {
                                        doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                    }

                                    doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), invtg[i]["LotNo"].ToString());             //批号
                                    doc.ReplaceText(string.Format("[ExpDate{0:00}]", i + 1), invtg[i]["ExpDate"].ToString());         //有效日期
                                    doc.ReplaceText(string.Format("[RiDate{0:00}]", i + 1), invtg[i]["RiDate"].ToString());           //复检日期
                                    doc.ReplaceText(string.Format("[SourceDoc{0:00}]", i + 1), invtg[i]["SourceDoc"].ToString());     //来源单号
                                    doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), invtg[i]["Project"].ToString());         //项目代号

                                    string tempRemark = invtg[i]["Remark"].ToString();                                                //备注
                                    if (CurrentCompany == 2)
                                    {
                                        if (tempRemark.Length > 10)
                                        {
                                            tempRemark = BaseHelper.StrLeft(tempRemark, 9) + "...";
                                        }
                                    }
                                    else {
                                        if (tempRemark.Length > 17)
                                        {
                                            tempRemark = BaseHelper.StrLeft(tempRemark, 16) + "...";
                                        }
                                    }

                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                }

                                #region //剩餘欄位
                                int CompanydocRow = 0;
                                if (CurrentCompany == 2) {
                                    CompanydocRow = 12;
                                }
                                else {
                                    CompanydocRow = 9;
                                }
                                if (invtg.Count() < CompanydocRow)
                                {
                                    for (int i = invtg.Count(); i < CompanydocRow; i++)
                                    {
                                        doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                        if (i == invtg.Count() && i < CompanydocRow)
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                        }
                                        else
                                        {
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                        }

                                        doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[OutStock{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[InStock{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ReturnDate{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[PsQty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[GpQty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[RQty{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Muom{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[ExpDate{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[RiDate{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[SourceDoc{0:00}]", i + 1), "");
                                        doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), "");
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
                            wordFileName = string.Format(wordFileName, invtf[0]["TF001"].ToString(), invtf[0]["TF002"].ToString());

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
                                    doc.ReplaceText("[MQ002]", invtf[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TF001Doc]", invtf[0]["TF001Doc"].ToString());
                                    doc.ReplaceText("[TF002]", invtf[0]["TF002"].ToString());
                                    doc.ReplaceText("[TF024Date]", invtf[0]["TF024Date"].ToString());
                                    doc.ReplaceText("[TF004]", invtf[0]["TF004"].ToString());
                                    doc.ReplaceText("[TF015]", invtf[0]["TF015"].ToString());
                                    doc.ReplaceText("[TF016]", invtf[0]["TF016"].ToString());
                                    doc.ReplaceText("[TF018]", invtf[0]["TF018"].ToString());
                                    doc.ReplaceText("[TF007Dep]", invtf[0]["TF007Dep"].ToString());
                                    doc.ReplaceText("[TF008Staff]", invtf[0]["TF008Staff"].ToString());
                                    doc.ReplaceText("[TF009Fac]", invtf[0]["TF009Fac"].ToString());
                                    doc.ReplaceText("[TF005]", invtf[0]["TF005"].ToString());
                                    doc.ReplaceText("[TF011]", invtf[0]["TF011"].ToString());
                                    if (CurrentCompany == 2)
                                    {
                                        switch (invtf[0]["TF010"].ToString())
                                        {
                                            case "外加":
                                                doc.ReplaceText("[TF010]", "應稅外加");               //课税别
                                                break;
                                            case "內含":
                                                doc.ReplaceText("[TF010]", "應稅內含");               //课税别
                                                break;
                                            default:
                                                doc.ReplaceText("[TF010]", invtf[0]["TF010"].ToString());
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        doc.ReplaceText("[TF010]", invtf[0]["TF010"].ToString());               //课税别
                                    }
                                    doc.ReplaceText("[TF013]", invtf[0]["TF013"].ToString());
                                    doc.ReplaceText("[TF014]", invtf[0]["TF014"].ToString());
                                    doc.ReplaceText("[TF012Rate]", invtf[0]["TF012Rate"].ToString());
                                    doc.ReplaceText("[TF026]", invtf[0]["TF026"].ToString());
                                    doc.ReplaceText("[TF020]", invtf[0]["TF020"].ToString());


                                    doc.ReplaceText("[TF022Qua]", "************");
                                    doc.ReplaceText("[TF023Amount]", "************.**");
                                    doc.ReplaceText("[TF027Tax]", "************.**");
                                    doc.ReplaceText("[TotalAmount]", "************.**");

                                    DateTime currentDate = DateTime.Now;
                                    string formattedDate = currentDate.ToString("yyyy/MM/dd");
                                    doc.ReplaceText("[CreateDate]", formattedDate);     //製表日期

                                    doc.ReplaceText("[cp]", p.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());
                                    doc.ReplaceText("[NextPage]", "接下頁...");
                                    #endregion

                                    #region //單身
                                    int CompanydocRow = 0;
                                    if (CurrentCompany == 2)
                                    {
                                        CompanydocRow = 12;
                                    }
                                    else
                                    {
                                        CompanydocRow = 9;
                                    }

                                    if (invtg.Count() > 0)
                                    {
                                        for (int i = 0; i < CompanydocRow; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["MtlItemNo"].ToString());

                                            string tempMtlItemName = invtg[i + (p - 1) * CompanydocRow]["MtlItemName"].ToString();

                                            if (CurrentCompany == 2)
                                            {
                                                if (tempMtlItemName.Length > 26)
                                                {
                                                    tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 25) + "...";
                                                }
                                            }
                                            else
                                            {
                                                if (tempMtlItemName.Length > 20)
                                                {
                                                    tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 20) + "...";
                                                }
                                            }

                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[OutStock{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["OutStock"].ToString());
                                            doc.ReplaceText(string.Format("[InStock{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["InStock"].ToString());
                                            doc.ReplaceText(string.Format("[ReturnDate{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["ReturnDate"].ToString());
                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Qty"].ToString());
                                            doc.ReplaceText(string.Format("[PsQty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["PsQty"].ToString());
                                            doc.ReplaceText(string.Format("[GpQty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["GpQty"].ToString());
                                            doc.ReplaceText(string.Format("[RQty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["RQty"].ToString());
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Uom"].ToString());
                                            doc.ReplaceText(string.Format("[Muom{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Muom"].ToString());
                                            

                                            if (CurrentCompany == 2)
                                            {
                                                switch (invtg[i]["Closeout"].ToString())
                                                {
                                                    case "N":                                                        
                                                        doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "未結案");         //结案码
                                                        break;
                                                    case "Y":
                                                        doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "已結案");         //结案码
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Closeout"].ToString());    //结案码

                                            }

                                            if (showPrice)
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["UnPri"].ToString());
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), Convert.ToDouble(invtg[i + (p - 1) * CompanydocRow]["Amount"]).ToString("n0"));
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["LotNo"].ToString());
                                            doc.ReplaceText(string.Format("[ExpDate{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["ExpDate"].ToString());
                                            doc.ReplaceText(string.Format("[RiDate{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["RiDate"].ToString());
                                            doc.ReplaceText(string.Format("[SourceDoc{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["SourceDoc"].ToString());
                                            doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Project"].ToString());

                                            string tempRemark = invtg[i + (p - 1) * CompanydocRow]["Remark"].ToString();

                                            if (CurrentCompany == 2)
                                            {
                                                if (tempRemark.Length > 10)
                                                {
                                                    tempRemark = BaseHelper.StrLeft(tempRemark, 9) + "...";
                                                }
                                            }
                                            else
                                            {
                                                if (tempRemark.Length > 17)
                                                {
                                                    tempRemark = BaseHelper.StrLeft(tempRemark, 16) + "...";
                                                }
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
                                    doc.ReplaceText("[MQ002]", invtf[0]["MQ002"].ToString());
                                    doc.ReplaceText("[TF001Doc]", invtf[0]["TF001Doc"].ToString());
                                    doc.ReplaceText("[TF002]", invtf[0]["TF002"].ToString());
                                    doc.ReplaceText("[TF024Date]", invtf[0]["TF024Date"].ToString());
                                    doc.ReplaceText("[TF004]", invtf[0]["TF004"].ToString());
                                    doc.ReplaceText("[TF015]", invtf[0]["TF015"].ToString());
                                    doc.ReplaceText("[TF016]", invtf[0]["TF016"].ToString());
                                    doc.ReplaceText("[TF018]", invtf[0]["TF018"].ToString());
                                    doc.ReplaceText("[TF007Dep]", invtf[0]["TF007Dep"].ToString());
                                    doc.ReplaceText("[TF008Staff]", invtf[0]["TF008Staff"].ToString());
                                    doc.ReplaceText("[TF009Fac]", invtf[0]["TF009Fac"].ToString());
                                    doc.ReplaceText("[TF005]", invtf[0]["TF005"].ToString());
                                    doc.ReplaceText("[TF011]", invtf[0]["TF011"].ToString());
                                    if (CurrentCompany == 2)
                                    {
                                        switch (invtf[0]["TF010"].ToString())
                                        {
                                            case "外加":
                                                doc.ReplaceText("[TF010]", "應稅外加");               //课税别
                                                break;
                                            case "內含":
                                                doc.ReplaceText("[TF010]", "應稅內含");               //课税别
                                                break;
                                            default:
                                                doc.ReplaceText("[TF010]", invtf[0]["TF010"].ToString());
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        doc.ReplaceText("[TF010]", invtf[0]["TF010"].ToString());               //课税别
                                    }
                                    doc.ReplaceText("[TF013]", invtf[0]["TF013"].ToString());
                                    doc.ReplaceText("[TF014]", invtf[0]["TF014"].ToString());
                                    doc.ReplaceText("[TF012Rate]", invtf[0]["TF012Rate"].ToString());
                                    doc.ReplaceText("[TF026]", invtf[0]["TF026"].ToString());
                                    doc.ReplaceText("[TF020]", invtf[0]["TF020"].ToString());

                                    doc.ReplaceText("[cp]", totalPage.ToString());
                                    doc.ReplaceText("[tp]", totalPage.ToString());

                                    DateTime currentDate = DateTime.Now;
                                    string formattedDate = currentDate.ToString("yyyy/MM/dd");
                                    doc.ReplaceText("[CreateDate]", formattedDate);     //製表日期

                                    doc.ReplaceText("[TF022Qua]", invtf[0]["TF022Qua"].ToString());

                                    if (showPrice)
                                    {
                                        doc.ReplaceText("[TF023Amount]", invtf[0]["TF023Amount"].ToString());
                                        doc.ReplaceText("[TF027Tax]", invtf[0]["TF027Tax"].ToString());
                                        doc.ReplaceText("[TotalAmount]", invtf[0]["TotalAmount"].ToString());
                                    }
                                    else
                                    {
                                        doc.ReplaceText("[TF023Amount]", "");
                                        doc.ReplaceText("[TF027Tax]", "");
                                        doc.ReplaceText("[TotalAmount]", "");
                                    }
                                    #endregion

                                    #region //單身
                                    int CompanydocRow = 0;
                                    if (CurrentCompany == 2)
                                    {
                                        CompanydocRow = 12;
                                    }
                                    else
                                    {
                                        CompanydocRow = 9;
                                    }
                                    if (invtg.Count() > 0)
                                    {
                                        for (int i = 0; i < (invtg.Count() % CompanydocRow != 0 ? invtg.Count() % CompanydocRow : CompanydocRow); i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["MtlItemNo"].ToString());

                                            string tempMtlItemName = invtg[i + (p - 1) * CompanydocRow]["MtlItemName"].ToString();

                                            if (tempMtlItemName.Length > 25)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 25) + "...";
                                            }

                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["MtlItemSpec"].ToString());

                                            doc.ReplaceText(string.Format("[OutStock{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["OutStock"].ToString());
                                            doc.ReplaceText(string.Format("[InStock{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["InStock"].ToString());
                                            doc.ReplaceText(string.Format("[ReturnDate{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["ReturnDate"].ToString());
                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Qty"].ToString());
                                            doc.ReplaceText(string.Format("[PsQty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["PsQty"].ToString());
                                            doc.ReplaceText(string.Format("[GpQty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["GpQty"].ToString());
                                            doc.ReplaceText(string.Format("[RQty{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["RQty"].ToString());
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Uom"].ToString());
                                            doc.ReplaceText(string.Format("[Muom{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Muom"].ToString());
                                            
                                            if (CurrentCompany == 2)
                                            {
                                                switch (invtg[i]["Closeout"].ToString())
                                                {
                                                    case "N":
                                                        doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "未結案");         //结案码
                                                        break;
                                                    case "Y":
                                                        doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "已結案");         //结案码
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["Closeout"].ToString());    //结案码
                                            }


                                            if (showPrice)
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), invtg[i + (p - 1) * CompanydocRow]["UnPri"].ToString());
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), Convert.ToDouble(invtg[i + (p - 1) * CompanydocRow]["Amount"]).ToString("n0"));
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[UnPri{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), invtg[i + (p - 1) * 9]["LotNo"].ToString());
                                            doc.ReplaceText(string.Format("[ExpDate{0:00}]", i + 1), invtg[i + (p - 1) * 9]["ExpDate"].ToString());
                                            doc.ReplaceText(string.Format("[RiDate{0:00}]", i + 1), invtg[i + (p - 1) * 9]["RiDate"].ToString());
                                            doc.ReplaceText(string.Format("[SourceDoc{0:00}]", i + 1), invtg[i + (p - 1) * 9]["SourceDoc"].ToString());
                                            doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), invtg[i + (p - 1) * 9]["Project"].ToString());

                                            string tempRemark = invtg[i + (p - 1) * CompanydocRow]["Remark"].ToString();
                                            if (tempRemark.Length > 17)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 16) + "...";
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
                                    if (invtg.Count() % CompanydocRow != 0)
                                    {
                                        for (int i = invtg.Count() % CompanydocRow; i < CompanydocRow; i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                            if (i == invtg.Count() % CompanydocRow && i < CompanydocRow)
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                            }
                                            else
                                            {
                                                doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                            }

                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[OutStock{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[InStock{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ReturnDate{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[PsQty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[GpQty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[RQty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Muom{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Closeout{0:00}]", i + 1), "");

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

                                            doc.ReplaceText(string.Format("[LotNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ExpDate{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[RiDate{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[SourceDoc{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Project{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");
                                        }
                                    }
                                    #endregion

                                    #endregion
                                }
                            }

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

        #region //Excel
        #region //InventoryAgingReportExcelDownload
        public void InventoryAgingReportExcelDownload(string MtlItemNo = "", string AgingDate = "", string ChangeType = "", string DataDate = "")
        {
            try
            {
                WebApiLoginCheck("InventoryAgingReport", "read");

                #region //Request
                scmInventoryDA = new ScmInventoryDA();
                dataRequest = scmInventoryDA.GetInventoryAgingReport(MtlItemNo, AgingDate, ChangeType, DataDate
                    , 1, 9999);
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
                    headerStyle.Fill.BackgroundColor = XLColor.AppleGreen;
                    headerStyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
                    headerStyle.Font.FontSize = 14;
                    headerStyle.Font.Bold = true;
                    #endregion

                    #region //dateStyle
                    var dateStyle = XLWorkbook.DefaultStyle;
                    dateStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
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
                    string excelFileName = "庫齡資料Excel檔";
                    string excelsheetName = "庫齡清單詳細資料";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] headerValues = new string[] { "舊品號", "舊品名", "舊規格", "新品號", "新品名", "新規格", "單位", "庫別", "30天以內庫存數量", "30天以內庫存金額", "30天以內包裝數量", "30~90天庫存數量", "30~90天庫存金額", "30~90天包裝數量", "90~180天庫存數量", "90~180天庫存金額", "90~180天包裝數量", "	180~270天庫存數量", "180~270天庫存金額", "180~270天包裝數量", "270~365天庫存數量", "270~365天庫存金額", "270~365天包裝數量", "365天以上庫存數量", "365天以上庫存金額", "365天以上包裝數量" };
                    string colIndexValue = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 15;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 4)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, 8).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < headerValues.Length; i++)
                        {
                            colIndexValue = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndexValue).Value = headerValues[i];
                            worksheet.Cell(colIndexValue).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY
                        foreach (var item in data)
                        {
                            rowIndex++;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.LA001.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.MB002.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.MB003.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.NewMtlItemNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.NewMtlItemName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.NewMtlItemSpec.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.MB004.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.LA009 != null ? item.LA009.ToString() : "" + "-" + item.MC002 != null ? item.MC002.ToString() : "";
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.INV_QTY_30D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.INV_AMT_30D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = item.PKG_QTY_30D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_QTY_30_90D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_AMT_30_90D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PKG_QTY_30_90D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_QTY_90_180D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_AMT_90_180D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PKG_QTY_90_180D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_QTY_180_270D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_AMT_180_270D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PKG_QTY_180_270D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_QTY_270_365D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_AMT_270_365D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PKG_QTY_270_365D.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_QTY_365D_PLUS.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.INV_AMT_365D_PLUS.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).Value = item.PKG_QTY_365D_PLUS.ToString();
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        #endregion

                        #region //置中
                        worksheet.Columns("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Columns("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        //worksheet.Columns("A:C").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Columns("1:3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        //worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        //worksheet.Column("4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        #endregion

                        #region //保護及鎖定
                        //workbook.Protect("123");

                        //worksheet.Columns("A:E").Style.Protection.SetLocked(true);
                        worksheet.Columns(BaseHelper.NumberToChar(1) + ":" + BaseHelper.NumberToChar(5)).Style.Protection.SetLocked(true);
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
                        string conditionalFormatRange = BaseHelper.NumberToChar(7) + ":" + BaseHelper.NumberToChar(7);
                        foreach (var range in worksheet.Ranges(conditionalFormatRange))
                        {
                            range.AddConditionalFormat().WhenEquals("M")
                                //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                                .Font.SetFontColor(XLColor.BabyBlue);

                            range.AddConditionalFormat().WhenEquals("F")
                                //.Fill.SetBackgroundColor(XLColor.BabyBlue)
                                .Font.SetFontColor(XLColor.Red);
                        }

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
        #endregion
    }
}