using Helpers;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Business_Manager.Controllers
{
    public class RmaController : WebController
    {
        private RmaDA rmaDA = new RmaDA();

        #region //View
        public ActionResult ReturnMerchandiseAuthorization()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetRandomRmaNo 取得隨機退貨單號
        [HttpPost]
        public void GetRandomRmaNo()
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "add,update");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.GetRandomRmaNo();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetReturnMerchandiseAuthorization 取得退(換)貨資料
        [HttpPost]
        public void GetReturnMerchandiseAuthorization(int RmaId = -1, string RmaNo = "", string ErpNo = "", int CustomerId = -1, string SearchKey = ""
            , string ConfirmStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "read");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.GetReturnMerchandiseAuthorization(RmaId, RmaNo, ErpNo, CustomerId, SearchKey
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

        #region //GetRmaDetail 取得退(換)貨項目資料
        [HttpPost]
        public void GetRmaDetail(int RmaDetailId = -1, int RmaId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "read");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.GetRmaDetail(RmaDetailId, RmaId
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
        public void GetTempShippingNote(int RmaId = -1, int CustomerId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "read");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.GetTempShippingNote(RmaId, CustomerId, SearchKey
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

        #region //GetRmaTransferLog 取得退(換)貨拋轉單據紀錄資料
        [HttpPost]
        public void GetRmaTransferLog(int RmaId = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "read");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.GetRmaTransferLog(RmaId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddReturnMerchandiseAuthorization 退(換)貨資料新增
        [HttpPost]
        public void AddReturnMerchandiseAuthorization(string RmaNo = "", int CustomerId = -1, string RmaDate = "", string RmaType = ""
            , string RmaRemark = "", string DocType = "")
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "add");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.AddReturnMerchandiseAuthorization(RmaNo, CustomerId, RmaDate, RmaType
                    , RmaRemark, DocType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRmaDetail 退(換)貨項目新增
        [HttpPost]
        public void AddRmaDetail(int RmaId = -1, int MtlItemId = -1, string ItemName = ""
            , int RmaQty = -1, string ItemType = "", float FreebieOrSpareQty = -1, string RmaDesc = "", int TargetInventory = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "add");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.AddRmaDetail(RmaId, MtlItemId, ItemName
                    , RmaQty, ItemType, FreebieOrSpareQty, RmaDesc, TargetInventory);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRmaDetailByTsn 退(換)貨項目【暫出單】新增
        [HttpPost]
        public void AddRmaDetailByTsn(int RmaId = -1, string TsnDetail = "")
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "add");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.AddRmaDetailByTsn(RmaId, TsnDetail);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateReturnMerchandiseAuthorization 退(換)貨資料更新
        [HttpPost]
        public void UpdateReturnMerchandiseAuthorization(int RmaId = -1, int CustomerId = -1, string RmaDate = ""
            , string RmaType = "", string RmaRemark = "", string DocType = "", string TransferMode = "")
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "update");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.UpdateReturnMerchandiseAuthorization(RmaId, CustomerId, RmaDate
                    , RmaType, RmaRemark, DocType, TransferMode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReturnMerchandiseAuthorizationConfirm 退(換)貨資料確認
        [HttpPost]
        public void UpdateReturnMerchandiseAuthorizationConfirm(int RmaId = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "confirm");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.UpdateReturnMerchandiseAuthorizationConfirm(RmaId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateReturnMerchandiseAuthorizationReConfirm 退(換)貨資料反確認
        [HttpPost]
        public void UpdateReturnMerchandiseAuthorizationReConfirm(int RmaId = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "reconfirm");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.UpdateReturnMerchandiseAuthorizationReConfirm(RmaId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRmaDetail 退(換)貨項目資料更新
        [HttpPost]
        public void UpdateRmaDetail(int RmaDetailId = -1, int MtlItemId = -1, string ItemName = ""
            , int RmaQty = -1, float FreebieOrSpareQty = -1, string RmaDesc = "", int TargetInventory = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "update");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.UpdateRmaDetail(RmaDetailId, MtlItemId, ItemName
                    , RmaQty, FreebieOrSpareQty, RmaDesc, TargetInventory);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteReturnMerchandiseAuthorization 退(換)貨資料刪除
        [HttpPost]
        public void DeleteReturnMerchandiseAuthorization(int RmaId = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "delete");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.DeleteReturnMerchandiseAuthorization(RmaId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRmaDetail 退(換)貨項目資料刪除
        [HttpPost]
        public void DeleteRmaDetail(int RmaDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("ReturnMerchandiseAuthorization", "delete");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.DeleteRmaDetail(RmaDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //ReturnMerchandiseAuthorizationDocDownload 退(換)貨資料單據下載
        public void ReturnMerchandiseAuthorizationDocDownload(int RmaId = -1, string Doc = "")
        {
            try
            {
                if (!Regex.IsMatch(Doc, "^(rma|temp-shipping-return|inventory-transaction-1106|inventory-transaction-1109)$", RegexOptions.IgnoreCase)) throw new SystemException("【單據種類】錯誤!");

                #region //Request
                rmaDA = new RmaDA();
                dataRequest = rmaDA.GetReturnMerchandiseAuthorizationDoc(RmaId, Doc);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = "", user = ""
                        , password = BaseHelper.RandomCode(10);
                    int totalPage = 0;

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("單據資料不完整!");

                    switch (Doc)
                    {
                        case "rma":
                            wordFileName = "客戶退(換)貨通知單-{0}";
                            filePath = "~/WordTemplate/SCM/P-S001表11-02 客戶退(換)貨通知單.docx";
                            var rma = result["data"];

                            #region //產生文件
                            using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                            {
                                wordFileName = string.Format(wordFileName, rma[0]["RmaNo"].ToString());

                                #region //單頭
                                doc.ReplaceText("[RmaNo]", rma[0]["RmaNo"].ToString());
                                doc.ReplaceText("[Customer]", rma[0]["Customer"].ToString());
                                doc.ReplaceText("[RmaDate]", rma[0]["RmaDate"].ToString());
                                switch (rma[0]["RmaType"].ToString())
                                {
                                    case "1":
                                        doc.ReplaceText("[RmaType1]", "■");
                                        doc.ReplaceText("[RmaType2]", "□");
                                        doc.ReplaceText("[RmaType3]", "□");
                                        break;
                                    case "2":
                                        doc.ReplaceText("[RmaType1]", "□");
                                        doc.ReplaceText("[RmaType2]", "■");
                                        doc.ReplaceText("[RmaType3]", "□");
                                        break;
                                    case "3":
                                        doc.ReplaceText("[RmaType1]", "□");
                                        doc.ReplaceText("[RmaType2]", "□");
                                        doc.ReplaceText("[RmaType3]", "■");
                                        break;
                                    default:
                                        doc.ReplaceText("[RmaType1]", "□");
                                        doc.ReplaceText("[RmaType2]", "□");
                                        doc.ReplaceText("[RmaType3]", "□");
                                        break;
                                }
                                doc.ReplaceText("[RmaRemark]", rma[0]["RmaRemark"].ToString());
                                doc.ReplaceText("[ConfirmUser]", rma[0]["ConfirmUser"].ToString() + "_" + rma[0]["RmaDate"].ToString());
                                #endregion

                                #region //單身
                                if (rma[0]["RmaDetail"].ToString().Length > 0)
                                {
                                    var detail = JObject.Parse(rma[0]["RmaDetail"].ToString())["data"];

                                    for (int i = 0; i < detail.Count(); i++)
                                    {
                                        doc.ReplaceText(string.Format("[CustomerMtlItemNo{0:00}]", i + 1), detail[i]["CustomerMtlItemNo"].ToString());
                                        doc.ReplaceText(string.Format("[ItemName{0:00}]", i + 1), detail[i]["ItemName"].ToString());
                                        doc.ReplaceText(string.Format("[RmaQty{0:00}]", i + 1), detail[i]["RmaQty"].ToString());
                                        doc.ReplaceText(string.Format("[RmaDesc{0:00}]", i + 1), detail[i]["RmaDesc"].ToString());
                                    }

                                    if (detail.Count() < 10)
                                    {
                                        for (int i = detail.Count(); i < 10; i++)
                                        {
                                            doc.ReplaceText(string.Format("[CustomerMtlItemNo{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[ItemName{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[RmaQty{0:00}]", i + 1), "");
                                            doc.ReplaceText(string.Format("[RmaDesc{0:00}]", i + 1), "");
                                        }
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

                            fileGuid = Guid.NewGuid().ToString();
                            Session[fileGuid] = wordFile;
                            #endregion
                            break;
                        case "temp-shipping-return":
                            wordFileName = "暫出/入歸還單-{0}-{1}";
                            filePath = "~/WordTemplate/SCM/借出歸還單.docx";
                            secondFilePath = "~/WordTemplate/SCM/借出歸還單 換頁.docx";

                            var invth = result["data"];
                            var invti = result["dataDetail"];
                            user = result["user"].ToString();

                            #region //產生文件
                            totalPage = invti.Count() / 9 + (invti.Count() % 9 > 0 ? 1 : 0);

                            if (totalPage == 1)
                            {
                                #region //單頁
                                using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                                {
                                    wordFileName = string.Format(wordFileName, invth[0]["TH001"].ToString(), invth[0]["TH002"].ToString());

                                    #region //單頭
                                    doc.ReplaceText("[TH004]", invth[0]["TH004"].ToString());
                                    doc.ReplaceText("[TH011]", invth[0]["TH011"].ToString());
                                    doc.ReplaceText("[TH002]", invth[0]["TH002"].ToString());
                                    doc.ReplaceText("[TH005]", invth[0]["TH005"].ToString());
                                    doc.ReplaceText("[TH012]", invth[0]["TH012"].ToString());
                                    doc.ReplaceText("[TH023]", invth[0]["TH023"].ToString());
                                    doc.ReplaceText("[TH015]", invth[0]["TH015"].ToString());
                                    doc.ReplaceText("[TH025]", invth[0]["TH025"].ToString());
                                    doc.ReplaceText("[TH010]", invth[0]["TH010"].ToString());
                                    doc.ReplaceText("[TH007]", invth[0]["TH007"].ToString());
                                    doc.ReplaceText("[TH014]", invth[0]["TH014"].ToString());
                                    doc.ReplaceText("[TH008]", invth[0]["TH008"].ToString());

                                    doc.ReplaceText("[TH021]", invth[0]["TH021"].ToString());
                                    doc.ReplaceText("[TH026]", invth[0]["TH026"].ToString());
                                    doc.ReplaceText("[TH022]", invth[0]["TH022"].ToString());
                                    switch (invth[0]["TH010"].ToString())
                                    {
                                        case "1":
                                            doc.ReplaceText("[TsrnPrice]", (Convert.ToDouble(invth[0]["TH022"]) - Convert.ToDouble(invth[0]["TH026"])).ToString());
                                            break;
                                        default:
                                            doc.ReplaceText("[TsrnPrice]", invth[0]["TH022"].ToString());
                                            break;
                                    }

                                    doc.ReplaceText("[cp]", "1");
                                    doc.ReplaceText("[tp]", "1");
                                    doc.ReplaceText("[ConfirmUser]", user + "_" + invth[0]["TH023"].ToString());
                                    #endregion

                                    #region //單身
                                    if (invti.Count() > 0)
                                    {
                                        for (int i = 0; i < invti.Count(); i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invti[i]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invti[i]["MtlItemNo"].ToString());

                                            string tempMtlItemName = invti[i]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 33)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invti[i]["MtlItemSpec"].ToString());
                                            doc.ReplaceText(string.Format("[OutInvNo{0:00}]", i + 1), invti[i]["OutInvNo"].ToString());
                                            doc.ReplaceText(string.Format("[InInvNo{0:00}]", i + 1), invti[i]["InInvNo"].ToString());
                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invti[i]["Qty"].ToString());
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invti[i]["Uom"].ToString());
                                            doc.ReplaceText(string.Format("[UnitPrice{0:00}]", i + 1), invti[i]["UnitPrice"].ToString());
                                            doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), invti[i]["Amount"].ToString());
                                            doc.ReplaceText(string.Format("[SourceNote{0:00}]", i + 1), invti[i]["SourceNote"].ToString());

                                            string tempRemark = invti[i]["Remark"].ToString();
                                            if (tempRemark.Length > 13)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 13) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                        }

                                        #region //剩餘欄位
                                        if (invti.Count() < 9)
                                        {
                                            for (int i = invti.Count(); i < 9; i++)
                                            {
                                                doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                                if (i == invti.Count() && i < 9)
                                                {
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                                }
                                                else
                                                {
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                                }

                                                doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[OutInvNo{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[InInvNo{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[UnitPrice{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[SourceNote{0:00}]", i + 1), "");
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
                                    wordFileName = string.Format(wordFileName, invth[0]["TH001"].ToString(), invth[0]["TH002"].ToString());

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
                                            doc.ReplaceText("[TH004]", invth[0]["TH004"].ToString());
                                            doc.ReplaceText("[TH011]", invth[0]["TH011"].ToString());
                                            doc.ReplaceText("[TH002]", invth[0]["TH002"].ToString());
                                            doc.ReplaceText("[TH005]", invth[0]["TH005"].ToString());
                                            doc.ReplaceText("[TH012]", invth[0]["TH012"].ToString());
                                            doc.ReplaceText("[TH023]", invth[0]["TH023"].ToString());
                                            doc.ReplaceText("[TH015]", invth[0]["TH015"].ToString());
                                            doc.ReplaceText("[TH025]", invth[0]["TH025"].ToString());
                                            doc.ReplaceText("[TH010]", invth[0]["TH010"].ToString());
                                            doc.ReplaceText("[TH007]", invth[0]["TH007"].ToString());
                                            doc.ReplaceText("[TH014]", invth[0]["TH014"].ToString());
                                            doc.ReplaceText("[TH008]", invth[0]["TH008"].ToString());

                                            doc.ReplaceText("[TH021]", invth[0]["TH021"].ToString());
                                            doc.ReplaceText("[TH026]", invth[0]["TH026"].ToString());
                                            doc.ReplaceText("[TH022]", invth[0]["TH022"].ToString());
                                            switch (invth[0]["TH010"].ToString())
                                            {
                                                case "1":
                                                    doc.ReplaceText("[TsrnPrice]", (Convert.ToDouble(invth[0]["TH022"]) - Convert.ToDouble(invth[0]["TH026"])).ToString());
                                                    break;
                                                default:
                                                    doc.ReplaceText("[TsrnPrice]", invth[0]["TH022"].ToString());
                                                    break;
                                            }

                                            doc.ReplaceText("[cp]", p.ToString());
                                            doc.ReplaceText("[tp]", totalPage.ToString());
                                            doc.ReplaceText("[ConfirmUser]", user + "_" + invth[0]["TH023"].ToString());
                                            #endregion

                                            #region //單身
                                            if (invti.Count() > 0)
                                            {
                                                for (int i = 0; i < 9; i++)
                                                {
                                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invti[i + (p - 1) * 9]["Seq"].ToString());
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invti[i + (p - 1) * 9]["MtlItemNo"].ToString());

                                                    string tempMtlItemName = invti[i + (p - 1) * 9]["MtlItemName"].ToString();
                                                    if (tempMtlItemName.Length > 33)
                                                    {
                                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                                    }
                                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invti[i + (p - 1) * 9]["MtlItemSpec"].ToString());
                                                    doc.ReplaceText(string.Format("[OutInvNo{0:00}]", i + 1), invti[i + (p - 1) * 9]["OutInvNo"].ToString());
                                                    doc.ReplaceText(string.Format("[InInvNo{0:00}]", i + 1), invti[i + (p - 1) * 9]["InInvNo"].ToString());
                                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invti[i + (p - 1) * 9]["Qty"].ToString());
                                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invti[i + (p - 1) * 9]["Uom"].ToString());
                                                    doc.ReplaceText(string.Format("[UnitPrice{0:00}]", i + 1), invti[i + (p - 1) * 9]["UnitPrice"].ToString());
                                                    doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), invti[i + (p - 1) * 9]["Amount"].ToString());
                                                    doc.ReplaceText(string.Format("[SourceNote{0:00}]", i + 1), invti[i + (p - 1) * 9]["SourceNote"].ToString());

                                                    string tempRemark = invti[i + (p - 1) * 9]["Remark"].ToString();
                                                    if (tempRemark.Length > 13)
                                                    {
                                                        tempRemark = BaseHelper.StrLeft(tempRemark, 13) + "...";
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
                                            doc.ReplaceText("[TH004]", invth[0]["TH004"].ToString());
                                            doc.ReplaceText("[TH011]", invth[0]["TH011"].ToString());
                                            doc.ReplaceText("[TH002]", invth[0]["TH002"].ToString());
                                            doc.ReplaceText("[TH005]", invth[0]["TH005"].ToString());
                                            doc.ReplaceText("[TH012]", invth[0]["TH012"].ToString());
                                            doc.ReplaceText("[TH023]", invth[0]["TH023"].ToString());
                                            doc.ReplaceText("[TH015]", invth[0]["TH015"].ToString());
                                            doc.ReplaceText("[TH025]", invth[0]["TH025"].ToString());
                                            doc.ReplaceText("[TH010]", invth[0]["TH010"].ToString());
                                            doc.ReplaceText("[TH007]", invth[0]["TH007"].ToString());
                                            doc.ReplaceText("[TH014]", invth[0]["TH014"].ToString());
                                            doc.ReplaceText("[TH008]", invth[0]["TH008"].ToString());

                                            doc.ReplaceText("[cp]", totalPage.ToString());
                                            doc.ReplaceText("[tp]", totalPage.ToString());
                                            doc.ReplaceText("[ConfirmUser]", user + "_" + invth[0]["TH023"].ToString());
                                            #endregion

                                            #region //單身
                                            if (invti.Count() > 0)
                                            {
                                                for (int i = 0; i < (invti.Count() % 9 != 0 ? invti.Count() % 9 : 9); i++)
                                                {
                                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invti[i + (p - 1) * 9]["Seq"].ToString());
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invti[i + (p - 1) * 9]["MtlItemNo"].ToString());

                                                    string tempMtlItemName = invti[i + (p - 1) * 9]["MtlItemName"].ToString();
                                                    if (tempMtlItemName.Length > 33)
                                                    {
                                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 33) + "...";
                                                    }
                                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invti[i + (p - 1) * 9]["MtlItemSpec"].ToString());
                                                    doc.ReplaceText(string.Format("[OutInvNo{0:00}]", i + 1), invti[i + (p - 1) * 9]["OutInvNo"].ToString());
                                                    doc.ReplaceText(string.Format("[InInvNo{0:00}]", i + 1), invti[i + (p - 1) * 9]["InInvNo"].ToString());
                                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invti[i + (p - 1) * 9]["Qty"].ToString());
                                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invti[i + (p - 1) * 9]["Uom"].ToString());
                                                    doc.ReplaceText(string.Format("[UnitPrice{0:00}]", i + 1), invti[i + (p - 1) * 9]["UnitPrice"].ToString());
                                                    doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), invti[i + (p - 1) * 9]["Amount"].ToString());
                                                    doc.ReplaceText(string.Format("[SourceNote{0:00}]", i + 1), invti[i + (p - 1) * 9]["SourceNote"].ToString());

                                                    string tempRemark = invti[i + (p - 1) * 9]["Remark"].ToString();
                                                    if (tempRemark.Length > 13)
                                                    {
                                                        tempRemark = BaseHelper.StrLeft(tempRemark, 13) + "...";
                                                    }
                                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                                }
                                            }
                                            else
                                            {
                                                throw new SystemException("【單身資料】錯誤!");
                                            }

                                            #region //剩餘欄位
                                            if (invti.Count() % 9 != 0)
                                            {
                                                for (int i = invti.Count() % 9; i < 9; i++)
                                                {
                                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                                    if (i == invti.Count() % 9 && i < 9)
                                                    {
                                                        doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                                    }
                                                    else
                                                    {
                                                        doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                                    }

                                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[OutInvNo{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[InInvNo{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[UnitPrice{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[Amount{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[SourceNote{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), "");
                                                }
                                            }
                                            #endregion
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
                            break;
                        case "inventory-transaction-1106":
                        case "inventory-transaction-1109":
                            switch (Doc)
                            {
                                case "inventory-transaction-1106":
                                    wordFileName = "客供入料單-{0}-{1}";
                                    filePath = "~/WordTemplate/SCM/P-W001表10-01 客供入料單.docx";
                                    secondFilePath = "~/WordTemplate/SCM/P-W001表10-01 客供入料單 換頁.docx";
                                    break;
                                case "inventory-transaction-1109":
                                    wordFileName = "0成本入庫單-{0}-{1}";
                                    filePath = "~/WordTemplate/SCM/P-W001表12-01 0成本入庫單.docx";
                                    secondFilePath = "~/WordTemplate/SCM/P-W001表12-01 0成本入庫單 換頁.docx";
                                    break;
                            }

                            var invta = result["data"];
                            var invtb = result["dataDetail"];
                            user = result["user"].ToString();

                            #region //產生文件
                            totalPage = invtb.Count() / 4 + (invtb.Count() % 4 > 0 ? 1 : 0);

                            if (totalPage == 1)
                            {
                                #region //單頁
                                using (DocX doc = DocX.Load(Server.MapPath(filePath)))
                                {
                                    wordFileName = string.Format(wordFileName, invta[0]["TA001"].ToString(), invta[0]["TA002"].ToString());

                                    #region //單頭
                                    doc.ReplaceText("[TA002]", invta[0]["TA002"].ToString());
                                    doc.ReplaceText("[TA004]", invta[0]["TA004"].ToString());
                                    doc.ReplaceText("[TA005]", invta[0]["TA005"].ToString());
                                    doc.ReplaceText("[TA006]", invta[0]["TA006"].ToString());
                                    doc.ReplaceText("[TA010]", invta[0]["TA010"].ToString());
                                    doc.ReplaceText("[TA011]", invta[0]["TA011"].ToString());
                                    doc.ReplaceText("[TA014]", invta[0]["TA014"].ToString());

                                    doc.ReplaceText("[cp]", "1");
                                    doc.ReplaceText("[tp]", "1");
                                    doc.ReplaceText("[ConfirmUser]", (user + "_" + invta[0]["TA014"].ToString()));
                                    #endregion

                                    #region //單身
                                    if (invtb.Count() > 0)
                                    {
                                        for (int i = 0; i < invtb.Count(); i++)
                                        {
                                            doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invtb[i]["Seq"].ToString());
                                            doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invtb[i]["MtlItemNo"].ToString());

                                            string tempMtlItemName = invtb[i]["MtlItemName"].ToString();
                                            if (tempMtlItemName.Length > 28)
                                            {
                                                tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 28) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                            doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invtb[i]["MtlItemSpec"].ToString());
                                            doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), invtb[i]["InvNo"].ToString());
                                            doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), invtb[i]["InvName"].ToString());
                                            doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invtb[i]["Uom"].ToString());
                                            doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invtb[i]["Qty"].ToString());

                                            string tempRemark = invtb[i]["Remark"].ToString();
                                            if (tempRemark.Length > 13)
                                            {
                                                tempRemark = BaseHelper.StrLeft(tempRemark, 13) + "...";
                                            }
                                            doc.ReplaceText(string.Format("[Remark{0:00}]", i + 1), tempRemark);
                                        }

                                        #region //剩餘欄位
                                        if (invtb.Count() < 4)
                                        {
                                            for (int i = invtb.Count(); i < 4; i++)
                                            {
                                                doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                                if (i == invtb.Count() && i < 4)
                                                {
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                                }
                                                else
                                                {
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                                }

                                                doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                                doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
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
                                    wordFileName = string.Format(wordFileName, invta[0]["TA001"].ToString(), invta[0]["TA002"].ToString());

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
                                            doc.ReplaceText("[TA002]", invta[0]["TA002"].ToString());
                                            doc.ReplaceText("[TA004]", invta[0]["TA004"].ToString());
                                            doc.ReplaceText("[TA005]", invta[0]["TA005"].ToString());
                                            doc.ReplaceText("[TA006]", invta[0]["TA006"].ToString());
                                            doc.ReplaceText("[TA010]", invta[0]["TA010"].ToString());
                                            doc.ReplaceText("[TA014]", invta[0]["TA014"].ToString());

                                            doc.ReplaceText("[cp]", p.ToString());
                                            doc.ReplaceText("[tp]", totalPage.ToString());
                                            #endregion

                                            #region //單身
                                            if (invtb.Count() > 0)
                                            {
                                                for (int i = 0; i < 4; i++)
                                                {
                                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invtb[i + (p - 1) * 4]["Seq"].ToString());
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invtb[i + (p - 1) * 4]["MtlItemNo"].ToString());

                                                    string tempMtlItemName = invtb[i + (p - 1) * 4]["MtlItemName"].ToString();
                                                    if (tempMtlItemName.Length > 28)
                                                    {
                                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 28) + "...";
                                                    }
                                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invtb[i + (p - 1) * 4]["MtlItemSpec"].ToString());
                                                    doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), invtb[i + (p - 1) * 4]["InvNo"].ToString());
                                                    doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), invtb[i + (p - 1) * 4]["InvName"].ToString());
                                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invtb[i + (p - 1) * 4]["Uom"].ToString());
                                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invtb[i + (p - 1) * 4]["Qty"].ToString());

                                                    string tempRemark = invtb[i + (p - 1) * 4]["Remark"].ToString();
                                                    if (tempRemark.Length > 13)
                                                    {
                                                        tempRemark = BaseHelper.StrLeft(tempRemark, 13) + "...";
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
                                            doc.ReplaceText("[TA002]", invta[0]["TA002"].ToString());
                                            doc.ReplaceText("[TA004]", invta[0]["TA004"].ToString());
                                            doc.ReplaceText("[TA005]", invta[0]["TA005"].ToString());
                                            doc.ReplaceText("[TA006]", invta[0]["TA006"].ToString());
                                            doc.ReplaceText("[TA010]", invta[0]["TA010"].ToString());
                                            doc.ReplaceText("[TA011]", invta[0]["TA011"].ToString());
                                            doc.ReplaceText("[TA014]", invta[0]["TA014"].ToString());

                                            doc.ReplaceText("[cp]", totalPage.ToString());
                                            doc.ReplaceText("[tp]", totalPage.ToString());
                                            doc.ReplaceText("[ConfirmUser]", user + "_" + invta[0]["TA014"].ToString());
                                            #endregion

                                            #region //單身
                                            if (invtb.Count() > 0)
                                            {
                                                for (int i = 0; i < (invtb.Count() % 4 != 0 ? invtb.Count() % 4 : 4); i++)
                                                {
                                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), invtb[i + (p - 1) * 4]["Seq"].ToString());
                                                    doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), invtb[i + (p - 1) * 4]["MtlItemNo"].ToString());

                                                    string tempMtlItemName = invtb[i + (p - 1) * 4]["MtlItemName"].ToString();
                                                    if (tempMtlItemName.Length > 28)
                                                    {
                                                        tempMtlItemName = BaseHelper.StrLeft(tempMtlItemName, 28) + "...";
                                                    }
                                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), tempMtlItemName);

                                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), invtb[i + (p - 1) * 4]["MtlItemSpec"].ToString());
                                                    doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), invtb[i + (p - 1) * 4]["InvNo"].ToString());
                                                    doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), invtb[i + (p - 1) * 4]["InvName"].ToString());
                                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), invtb[i + (p - 1) * 4]["Uom"].ToString());
                                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), invtb[i + (p - 1) * 4]["Qty"].ToString());

                                                    string tempRemark = invtb[i + (p - 1) * 4]["Remark"].ToString();
                                                    if (tempRemark.Length > 13)
                                                    {
                                                        tempRemark = BaseHelper.StrLeft(tempRemark, 13) + "...";
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
                                            if (invtb.Count() % 4 != 0)
                                            {
                                                for (int i = invtb.Count() % 4; i < 4; i++)
                                                {
                                                    doc.ReplaceText(string.Format("[Seq{0:00}]", i + 1), "");

                                                    if (i == invtb.Count() % 4 && i < 4)
                                                    {
                                                        doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "以下空白//");
                                                    }
                                                    else
                                                    {
                                                        doc.ReplaceText(string.Format("[MtlItemNo{0:00}]", i + 1), "");
                                                    }

                                                    doc.ReplaceText(string.Format("[MtlItemName{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[MtlItemSpec{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[InvNo{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[InvName{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[Uom{0:00}]", i + 1), "");
                                                    doc.ReplaceText(string.Format("[Qty{0:00}]", i + 1), "");
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
                            break;
                    }

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