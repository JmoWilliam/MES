using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Helpers;
using Newtonsoft.Json.Linq;
using SCMDA;
using ClosedXML.Excel;


namespace Business_Manager.Controllers
{
    public class LotNumberController : WebController
    {
        private LotNumberDA lotNumberDA = new LotNumberDA();

        #region //View
        public ActionResult LotNumberManagement()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetLotNumber 取得批號資料
        [HttpPost]
        public void GetLotNumber(int LotNumberId = -1, string MtlItemNo = "", string MtlItemName = "", string LotNumberNo = "", string CloseStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LotNumberManagement", "read,constrained-data");

                #region //Request
                lotNumberDA = new LotNumberDA();
                dataRequest = lotNumberDA.GetLotNumber(LotNumberId, MtlItemNo, MtlItemName, LotNumberNo, CloseStatus
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

        #region //GetLnDetail 取得批號詳細資料
        [HttpPost]
        public void GetLnDetail(int LnDetailId = -1, int LotNumberId = -1, string MtlItemNo = "", string MtlItemName = "", string FromErpFullNo = "", int InventoryId = -1, int TransactionType = -1, string DocType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LotNumberManagement", "read,constrained-data");

                #region //Request
                lotNumberDA = new LotNumberDA();
                dataRequest = lotNumberDA.GetLnDetail(LnDetailId, LotNumberId, MtlItemNo, MtlItemName, FromErpFullNo, InventoryId, TransactionType, DocType
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

        #region //Add
        #region //AddLotNumber 新增批號資料
        public void AddLotNumber(int MtlItemId = -1, string LotNumberNo = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("LotNumberManagement", "add");

                #region //Request
                lotNumberDA = new LotNumberDA();
                dataRequest = lotNumberDA.AddLotNumber(MtlItemId, LotNumberNo, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateLotNumber 更新批號資料
        public void UpdateLotNumber(int LotNumberId = -1, int MtlItemId = -1, string LotNumberNo = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("LotNumberManagement", "add");

                #region //Request
                lotNumberDA = new LotNumberDA();
                dataRequest = lotNumberDA.UpdateLotNumber(LotNumberId, MtlItemId, LotNumberNo, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLotNumberManualSynchronize 批號資料手動同步
        [HttpPost]
        public void UpdateLotNumberManualSynchronize(string MtlItemNo = "", string LotNumber = "", string SyncStartDate = "", string SyncEndDate = ""
            , string NormalSync = "", string TranSync = "")
        {
            try
            {
                WebApiLoginCheck("GoodsReceiptManagement", "sync");

                #region //Request
                lotNumberDA = new LotNumberDA();
                dataRequest = lotNumberDA.UpdateLotNumberManualSynchronize(MtlItemNo, LotNumber, SyncStartDate, SyncEndDate
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
        #endregion

        #region //Delete
        #region //DeleteLotNumber 批號資料刪除
        [HttpPost]
        public void DeleteLotNumber(int LotNumberId = -1)
        {
            try
            {
                WebApiLoginCheck("LotNumberManagement", "delete");

                #region //Request
                lotNumberDA = new LotNumberDA();
                dataRequest = lotNumberDA.DeleteLotNumber(LotNumberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateLotNumberSynchronize 批號資料同步
        [HttpPost]
        [Route("api/ERP/LotNumberSynchronize")]
        public void UpdateLotNumberSynchronize(string Company, string SecretKey, string UpdateDate, string NormalSync, string TranSync)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateLotNumberSynchronize");
                #endregion

                #region //Request
                dataRequest = lotNumberDA.UpdateLotNumberSynchronize(Company, UpdateDate, NormalSync, TranSync);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        //#region //IntoLotBindBarcode -- 帶入Lot BindBarode Excel
        //[HttpPost]
        //public void IntoLotBindBarcode(int FileId = -1, string CompanyNo = "")
        //{
        //    try
        //    {
        //        #region //解析EXCEL
        //        //FilePath = "\\\\192.168.20.199\\d8-系統開發室\\99-單位共用區\\安宏\\EXCEL_TEST0315.xlsx";
        //        XLWorkbook workbook = new XLWorkbook(FilePath);
        //        IXLWorksheet worksheet = workbook.Worksheet(1);
        //        var firstCell = worksheet.FirstCellUsed();
        //        var lastCell = worksheet.LastCellUsed();

        //        int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

        //        var data = worksheet.Range(firstCell.Address, lastCell.Address);
        //        var table = data.AsTable();

        //        List<LotBarcode> qcItems = new List<LotBarcode>();
        //        for (var i = 1; i <= CellLength; i++)
        //        {
        //            LotBarcode qcItem = new LotBarcode()
        //            {
        //                BarcodeNo = table.Cell(i, 2).Value.ToString(),
        //                BarcodeCount = Convert.ToInt32(table.Cell(i, 2).Value)
                        
        //            };
        //            qcItems.Add(qcItem);
        //        }
        //        #endregion

        //        #region //Request
        //        lotNumberDA = new LotNumberDA();
        //        dataRequest = lotNumberDA.IntoLotBindBarcode(qcItems, CompanyNo);
        //        JObject dataRequestJson = JObject.Parse(dataRequest);
        //        if (dataRequestJson["status"].ToString() != "success")
        //        {
        //            throw new SystemException(dataRequestJson["msg"].ToString());
        //        }
        //        #endregion

        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "success",
        //            msg = "OK",
        //        });
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
        //#endregion

        #endregion
    }
}