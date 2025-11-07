using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2013.Excel;
using Helpers;
using Newtonsoft.Json.Linq;
using QMSDA;
using SCMDA;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Xceed.Document.NET;
using ZXing;

namespace Business_Manager.Controllers
{
    public class QcDataManagementController : WebController
    {
        private QcDataManagementDA qcDataManagementDA = new QcDataManagementDA();

        private const string DataPath = @"~/TempQcData";
        private const string QcReportDataPath = @"~/QcReportDataPath";
        private const string ResolvesPath = @"~/ResolvesPath";

        #region //View
        public ActionResult QcDataManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QcGoodsReceiptManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QcPlanningKanban()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QcOutsourcingManagement()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetQcRecord 取得量測資料
        [HttpPost]
        public void GetQcRecord(int QcRecordId = -1, string WoErpFullNo = "", string QcNoticeNo = "", string MtlItemNo = "", string MtlItemName = ""
            , int QcTypeId = -1, string CheckQcMeasureData = "", int UserId = -1, string StartDate = "", string EndDate = "", int MoId = -1, int ModeId = -1, string QcType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcRecord(QcRecordId, WoErpFullNo, QcNoticeNo, MtlItemNo, MtlItemName
                    , QcTypeId, CheckQcMeasureData, UserId, StartDate, EndDate, MoId, ModeId, QcType
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

        #region //GetLetteringBarcodeData 取得刻字條碼
        [HttpPost]
        public void GetLetteringBarcodeData(int MoId = -1, int ItemSeq = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetLetteringBarcodeData(MoId, ItemSeq);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeAccurately 驗證條碼是否可以進行量測數據上傳
        [HttpPost]
        public void GetBarcodeAccurately(int QcRecordId = -1, string BarcodeListData = "", string QcType = "", int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetBarcodeAccurately(QcRecordId, BarcodeListData, QcType, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLotNumberAccurately 驗證批號是否可以進行量測數據上傳
        [HttpPost]
        public void GetLotNumberAccurately(int QcRecordId = -1, string LotNumberListData = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetLotNumberAccurately(QcRecordId, LotNumberListData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQmmDetail 取得量測機台資料
        [HttpPost]
        public void GetQmmDetail(int QmmDetailId = -1, int ShopId = -1, string MachineNumber = "", int QcMachineModeId = -1, string QcMachineModeList = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQmmDetail(QmmDetailId, ShopId, MachineNumber, QcMachineModeId, QcMachineModeList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetWorkShop 取得車間資料
        [HttpPost]
        public void GetWorkShop(int ShopId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetWorkShop(ShopId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeData 取得條碼資料
        [HttpPost]
        public void GetBarcodeData(int MoId = -1, string BarcodeNo = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetBarcodeData(MoId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMeasureData 取得量測資料
        [HttpPost]
        public void GetQcMeasureData(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMeasureData(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCheckQcMeasureData 取得量測單據狀態
        [HttpPost]
        public void GetCheckQcMeasureData(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetCheckQcMeasureData(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcItemPrinciple 取得量測項目編碼原則
        [HttpPost]
        public void GetQcItemPrinciple(int PrincipleId = -1, int QcClassId = -1, string PrincipleNo = "", string PrincipleDesc = "", int QmmDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcItemPrinciple(PrincipleId, QcClassId, PrincipleNo, PrincipleDesc, QmmDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcType 取得量測類型資料
        [HttpPost]
        public void GetQcType(int QcTypeId = -1, int ModeId = -1, string QcTypeNo = "", string QcTypeName = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcType(QcTypeId, ModeId, QcTypeNo, QcTypeName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetFileInfo 取得檔案資料
        [HttpPost]
        public void GetFileInfo(string FileIdList = "", string FileType = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdList, FileType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcClass 取得量測類型
        [HttpPost]
        public void GetQcClass(int QcClassId = -1, string QcClassNo = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcClass(QcClassId, QcClassNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcRecordFile 取得量測檔案資料
        [HttpPost]
        public void GetQcRecordFile(int QcRecordId = -1, string FileIdList = "", string FileType = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcRecordFile(QcRecordId, FileIdList, FileType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeInfo 取得條碼相關資料
        [HttpPost]
        public void GetBarcodeInfo(string BarcodeNo = "", string InputType = "", int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetBarcodeInfo(BarcodeNo, InputType, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcGoodsReceipt 取得進貨檢單據資料
        [HttpPost]
        public void GetQcGoodsReceipt(int QcGoodsReceiptId = -1, int QcRecordId = -1, string GrErpFullNo = "", string MtlItemNo = "", string MtlItemName = ""
            , string CheckQcMeasureData = "", int UserId = -1, string StartDate = "", string EndDate = "", string QcType = ""
            , string PoErpFullNo = "", int PoUserId = -1, string PrErpFullNo = "", int PrUserId = -1

            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcGoodsReceipt(QcGoodsReceiptId, QcRecordId, GrErpFullNo, MtlItemNo, MtlItemName
                    , CheckQcMeasureData, UserId, StartDate, EndDate, QcType
                    , PoErpFullNo, PoUserId, PrErpFullNo, PrUserId
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

        #region //GetQcRecordPlanning 取得量測單據排程資料
        [HttpPost]
        public void GetQcRecordPlanning(int QrpId = -1, int QcRecordId = -1, int QcMachineModeId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcRecordPlanning(QrpId, QcRecordId, QcMachineModeId
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

        #region //GetQcRecordQcItem 取得量測單據排程資料
        [HttpPost]
        public void GetQcRecordQcItem(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcRecordQcItem(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcItem 取得量測項目
        [HttpPost]
        public void GetQcItem(int QcItemId = -1, string QcItemNo = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcItem(QcItemId, QcItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMachineMode 取得量測機型、機台資料
        [HttpPost]
        public void GetQcMachineMode(int QmmDetailId = -1, string MachineNumber = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMachineMode(QmmDetailId, MachineNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMachineModeInfo 取得量測機型資料
        [HttpPost]
        public void GetQcMachineModeInfo(int QcMachineModeId = -1, string QcMachineModeNumber = "", string QcMachineModeNo = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMachineModeInfo(QcMachineModeId, QcMachineModeNumber, QcMachineModeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMachineModePlanning 取得量測機台排程資料
        [HttpPost]
        public void GetQcMachineModePlanning(int QmmpId = -1, int QrpId = -1, int QmmDetailId = -1, string StartDate = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMachineModePlanning(QmmpId, QrpId, QmmDetailId, StartDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetAllQcMachineMode 取得全部機型資料
        [HttpPost]
        public void GetAllQcMachineMode()
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetAllQcMachineMode();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLotNumber 取得批號資料
        [HttpPost]
        public void GetLotNumber(string MtlItemNo = "", string LotNumber = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetLotNumber(MtlItemNo, LotNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcWhitelist 取得上傳路徑白名單
        [HttpPost]
        public void GetQcWhitelist(int ListId = -1, int DepartmentId = -1, string FolderPath = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcWhitelist(ListId, DepartmentId, FolderPath);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMachineModePlanningForKanban 取得量測機台排程資料(For量測看板)
        [HttpPost]
        public void GetQcMachineModePlanningForKanban(string QcMachineModeList = "", string StartDate = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMachineModePlanningForKanban(QcMachineModeList, StartDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMeasureDataTemp 取得量測暫存資料
        [HttpPost]
        public void GetQcMeasureDataTemp(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMeasureDataTemp(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcMachineModePlanningSort 取得量測機台排程最大排序
        [HttpPost]
        public void GetQcMachineModePlanningSort(int QmmDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcMachineModePlanningSort(QmmDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcGoodsReceiptLog 取得進貨檢驗作業資料
        [HttpPost]
        public void GetQcGoodsReceiptLog(int LogId = -1, int QcGoodsReceiptId = -1)
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "gr-inspection");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcGoodsReceiptLog(LogId, QcGoodsReceiptId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetOutsourcingQcRecord 取得量測資料
        [HttpPost]
        public void GetOutsourcingQcRecord(int QcRecordId = -1, string QcNoticeNo = "", string MtlItemNo = "", string MtlItemName = "", int MtlItemId = -1
            , int QcTypeId = -1, string CheckQcMeasureData = "", int UserId = -1, string StartDate = "", string EndDate = "", string QcType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetOutsourcingQcRecord(QcRecordId, QcNoticeNo, MtlItemNo, MtlItemName, MtlItemId
                    , QcTypeId, CheckQcMeasureData, UserId, StartDate, EndDate, QcType
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

        #region //GetQcOutsourcingMeasureDataTemp 取得量測暫存資料
        [HttpPost]
        public void GetQcOutsourcingMeasureDataTemp(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "read");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcOutsourcingMeasureDataTemp(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//掃碼取得量測紀錄 
        [HttpGet]
        [Route("api/QMS/GetQcRecodDataQrcode")]
        public void GetQcRecodDataQrcode(int QcRecordId = -1)
        {
            try
            {
                if (QcRecordId <= 0)
                    throw new Exception("量測單號不能為空");
                string baseUrl = Request.Url.GetLeftPart(UriPartial.Authority);


                // 取得網址路徑
                qcDataManagementDA = new QcDataManagementDA();
                var resultJson = qcDataManagementDA.GetQcRecodDataQrcode(QcRecordId, baseUrl);
                JObject obj = JObject.Parse(resultJson);

                if (obj["status"]?.ToString() == "success")
                {
                    string qcPath = obj["data"]?.ToString();

                    if (!string.IsNullOrWhiteSpace(qcPath))
                    {
                        // 成功則直接導向網址
                        Response.Redirect(qcPath, endResponse: true);
                        return;
                    }
                    else
                    {
                        Response.Write("錯誤：查無對應網址");
                    }
                }
                else
                {
                    Response.Write("錯誤：" + (obj["msg"]?.ToString() ?? "未知錯誤"));
                }
            }
            catch (Exception e)
            {
                Response.Write("錯誤：" + e.Message);
                logger.Error(e);
            }
        }




        #region //GetQcRecodDataQrcode2 
        [HttpPost]
        public void GetQcRecodDataQrcode2(int QcRecordId = -1)
        {
            try
            {
                //WebApiLoginCheck("QcDataManagement", "read");
                string baseUrl = Request.Url.GetLeftPart(UriPartial.Authority);


                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcRecodDataQrcode(QcRecordId, baseUrl);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Add
        #region //AddQcRecord 新增量測紀錄單頭
        [HttpPost]
        public void AddQcRecord(int QcNoticeId = -1, int QcTypeId = -1, int MoId = -1, int MoProcessId = -1, int QmmDetailId = -1, int QcItemId = -1, string Remark = "", int CurrentFileId = -1, string SupportAqFlag = "", string LotNumber = ""
            , string QcRecordFile = "", string InputType = "", string ResolveFile = "", string ResolveFileJson = "", string QcRecordFileByNas = "", string QcMeasureDataJson = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "add");

                string ServerPath = Server.MapPath(DataPath);
                string ServerPath2 = Server.MapPath(QcReportDataPath);
                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddQcRecord(QcNoticeId, QcTypeId, MoId, MoProcessId, QmmDetailId, QcItemId, Remark, CurrentFileId, SupportAqFlag, LotNumber
                    , QcRecordFile, InputType, ServerPath, ServerPath2, ResolveFile, ResolveFileJson, QcRecordFileByNas, QcMeasureDataJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQcRecordFile 新增量測檔案(量測檔案歸檔功能)
        [HttpPost]
        public void AddQcRecordFile(int QcRecordId = -1, string InputType = "", string FileInfo = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "add");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddQcRecordFile(QcRecordId, InputType, FileInfo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQcGoodsReceipt 新增進貨檢量測單據
        [HttpPost]
        public void AddQcGoodsReceipt(int GrDetailId = -1, string InputType = "", string QcRecordFile = "", string ResolveFile = "", string ResolveFileJson = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "add");

                string ServerPath = Server.MapPath(DataPath);
                string ServerPath2 = Server.MapPath(QcReportDataPath);
                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddQcGoodsReceipt(GrDetailId, InputType, QcRecordFile, ResolveFile, ResolveFileJson, Remark, ServerPath, ServerPath2);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQcRecordPlanning 新增量測單據排程資料
        [HttpPost]
        public void AddQcRecordPlanning(int QcRecordId = -1, string UploadJson = "", string Spreadsheet = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "add");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddQcRecordPlanning(QcRecordId, UploadJson, Spreadsheet);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQcMachineModePlanning 新增量測機台排程資料
        [HttpPost]
        public void AddQcMachineModePlanning(int QrpId = -1, int QmmDetailId = -1, int Sort = -1)
        {
            try
            {
                //WebApiLoginCheck("QcDataManagement", "add");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddQcMachineModePlanning(QrpId, QmmDetailId, Sort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQcGoodsReceiptLog 新增進貨檢驗作業資料
        [HttpPost]
        public void AddQcGoodsReceiptLog(int QcGoodsReceiptId = -1, int ReceiptQty = -1, int AcceptQty = -1, int ReturnQty = -1
            , string AcceptanceDate = "", string QcStatus = "", string QuickStatus = "", string Remark = "", string QcGoodsReceiptLogFile = "")
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "gr-inspection");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddQcGoodsReceiptLog(QcGoodsReceiptId, ReceiptQty, AcceptQty, ReturnQty
                    , AcceptanceDate, QcStatus, QuickStatus, Remark, QcGoodsReceiptLogFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddOutsourcingQcRecord 新增量測紀錄單頭
        [HttpPost]
        public void AddOutsourcingQcRecord(int QcNoticeId = -1, int QcTypeId = -1, int MtlItemId = -1, int QmmDetailId = -1, int QcItemId = -1, string Remark = "", int CurrentFileId = -1, string SupportAqFlag = "", string LotNumber = ""
            , string QcRecordFile = "", string InputType = "", string ResolveFile = "", string ResolveFileJson = "", string QcRecordFileByNas = "", string QcMeasureDataJson = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "add");

                string ServerPath = Server.MapPath(DataPath);
                string ServerPath2 = Server.MapPath(QcReportDataPath);
                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddOutsourcingQcRecord(QcNoticeId, QcTypeId, MtlItemId, QmmDetailId, QcItemId, Remark, CurrentFileId, SupportAqFlag, LotNumber
                    , QcRecordFile, InputType, ServerPath, ServerPath2, ResolveFile, ResolveFileJson, QcRecordFileByNas, QcMeasureDataJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateQcRecord 更新量測記錄單頭資料
        [HttpPost]
        public void UpdateQcRecord(int QcRecordId = -1, int QcNoticeId = -1, int QcTypeId = -1, int MoId = -1, int MoProcessId = -1, string Remark = "", int CurrentFileId = -1, string QcRecordFile = "", string InputType = "", string SupportAqFlag = "", string ResolveFile = "", string ResolveFileJson = "", string QcRecordFileByNas = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                string ServerPath = Server.MapPath(DataPath);
                string ServerPath2 = Server.MapPath(QcReportDataPath);

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateQcRecord(QcRecordId, QcNoticeId, QcTypeId, MoId, MoProcessId, Remark, CurrentFileId, QcRecordFile, InputType, ServerPath, ServerPath2, SupportAqFlag, ResolveFile, ResolveFileJson, QcRecordFileByNas);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //VoidQcRecord 作廢量測記錄單頭資料
        [HttpPost]
        public void VoidQcRecord(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "void");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.VoidQcRecord(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateResolveFileJson 暫存量測檔案解析JSON
        [HttpPost]
        public void UpdateResolveFileJson(int QcRecordId = -1, string ResolveFileJson = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateResolveFileJson(QcRecordId, ResolveFileJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCheckGoods 確認收貨狀態
        [HttpPost]
        public void UpdateCheckGoods(int QcRecordId = -1, string CheckQcMeasureData = "", string Disallowance = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateCheckGoods(QcRecordId, CheckQcMeasureData, Disallowance);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUrgent 更改急件狀態
        [HttpPost]
        public void UpdateUrgent(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "urgent");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateUrgent(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQcGoodsReceipt 更新進貨檢資料
        [HttpPost]
        public void UpdateQcGoodsReceipt(int QcRecordId = -1, int GrDetailId = -1, string InputType = "", string QcRecordFile = "", string ResolveFile = "", string ResolveFileJson = "", string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                string ServerPath = Server.MapPath(DataPath);
                string ServerPath2 = Server.MapPath(QcReportDataPath);

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateQcGoodsReceipt(QcRecordId, GrDetailId, InputType, QcRecordFile, ResolveFile, ResolveFileJson, Remark, ServerPath, ServerPath2);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQcMachineModePlanningSort 更新量測機台排程順序
        [HttpPost]
        public void UpdateQcMachineModePlanningSort(string PlanningSort = "")
        {
            try
            {
                //WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateQcMachineModePlanningSort(PlanningSort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateEstimatedMeasurementTime 更新量測單據排程預計量測時間
        [HttpPost]
        public void UpdateEstimatedMeasurementTime(string UploadJson = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateEstimatedMeasurementTime(UploadJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePlanningSpreadsheetDate 更新重置排程暫存Spreadsheet
        [HttpPost]
        public void UpdatePlanningSpreadsheetDate(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdatePlanningSpreadsheetDate(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCheckGoodsByPlanning 確認收貨狀態(從排程啟動)
        [HttpPost]
        public void UpdateCheckGoodsByPlanning(int QcRecordId = -1, string CheckQcMeasureData = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateCheckGoodsByPlanning(QcRecordId, CheckQcMeasureData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePlanningConfirmStatus 更新計畫排程確認狀態
        [HttpPost]
        public void UpdatePlanningConfirmStatus(int QrpId = -1, string ConfirmStatus = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdatePlanningConfirmStatus(QrpId, ConfirmStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQcGoodsReceiptLog 更新進貨檢驗作業資料
        [HttpPost]
        public void UpdateQcGoodsReceiptLog(int LogId = -1, int QcGoodsReceiptId = -1, int ReceiptQty = -1, int AcceptQty = -1, int ReturnQty = -1
            , string AcceptanceDate = "", string QcStatus = "", string QuickStatus = "", string Remark = "", string QcGoodsReceiptLogFile = "")
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "gr-inspection");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateQcGoodsReceiptLog(LogId, QcGoodsReceiptId, ReceiptQty, AcceptQty, ReturnQty
                    , AcceptanceDate, QcStatus, QuickStatus, Remark, QcGoodsReceiptLogFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CancelMeasurementUpload 更改量測單據狀態
        [HttpPost]
        public void CancelMeasurementUpload(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.CancelMeasurementUpload(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ImportExcelToAddQcRecord 解析Excel新增QcRecord
        [HttpPost]
        public void ImportExcelToAddQcRecord(string FileId = "")
        {
            try
            {
                if (FileId.Length <= 0) throw new SystemException("必須上傳檔案!!");

                List<QcExcelFormat> qcExcelFormats = new List<QcExcelFormat>();

                string[] FileIds = FileId.Split(',');
                foreach (var fileId in FileIds)
                {
                    #region //取得File資料
                    BmFileInfo fileInfo = qcDataManagementDA.GetFileInfoById(Convert.ToInt32(fileId));
                    if (fileInfo == null)
                    {
                        throw new SystemException("查無此File資料!!");
                    }
                    #endregion

                    #region //解析EXCEL
                    using (var stream = new MemoryStream(fileInfo.FileContent))
                    {
                        using (XLWorkbook workbook = new XLWorkbook(stream))
                        {
                            IXLWorksheet worksheet = workbook.Worksheet(1);
                            var firstCell = worksheet.FirstCellUsed();
                            var lastCell = worksheet.LastCellUsed();

                            int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                            var data = worksheet.Range(firstCell.Address, lastCell.Address);
                            var table = data.AsTable();

                            for (var i = 2; i <= CellLength; i++)
                            {
                                try
                                {
                                    QcExcelFormat qcExcelFormat = new QcExcelFormat()
                                    {
                                        QcType = table.Cell(i, 1).Value.ToString(),
                                        InputType = table.Cell(i, 2).Value.ToString(),
                                        WoErpFullNo = table.Cell(i, 3).Value.ToString(),
                                        MtlItemNo = table.Cell(i, 4).Value.ToString(),
                                        Remark = table.Cell(i, 5).Value.ToString(),
                                        QcItemFormatPath = table.Cell(i, 6).Value.ToString(),
                                    };
                                    qcExcelFormats.Add(qcExcelFormat);
                                }
                                catch (Exception ex)
                                {
                                    throw new SystemException("第" + i + "筆資料ERROR:" + ex.Message);
                                }
                            }
                        }
                    }
                    #endregion
                }

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddBatchQcRecord(qcExcelFormats);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    jsonResponse = dataRequestJson;
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

        #region //ImportExcelToAddOutSourceQcRecord 解析Excel新增OutSourceQcRecord
        [HttpPost]
        public void ImportExcelToAddOutSourceQcRecord(string FileId = "")
        {
            try
            {
                if (FileId.Length <= 0) throw new SystemException("必須上傳檔案!!");

                List<QcExcelFormat> qcExcelFormats = new List<QcExcelFormat>();

                string[] FileIds = FileId.Split(',');
                foreach (var fileId in FileIds)
                {
                    #region //取得File資料
                    BmFileInfo fileInfo = qcDataManagementDA.GetFileInfoById(Convert.ToInt32(fileId));
                    if (fileInfo == null)
                    {
                        throw new SystemException("查無此File資料!!");
                    }
                    #endregion

                    #region //解析EXCEL
                    using (var stream = new MemoryStream(fileInfo.FileContent))
                    {
                        using (XLWorkbook workbook = new XLWorkbook(stream))
                        {
                            IXLWorksheet worksheet = workbook.Worksheet(1);
                            var firstCell = worksheet.FirstCellUsed();
                            var lastCell = worksheet.LastCellUsed();

                            int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                            var data = worksheet.Range(firstCell.Address, lastCell.Address);
                            var table = data.AsTable();

                            for (var i = 2; i <= CellLength; i++)
                            {
                                try
                                {
                                    QcExcelFormat qcExcelFormat = new QcExcelFormat()
                                    {
                                        QcType = table.Cell(i, 1).Value.ToString(),
                                        InputType = table.Cell(i, 2).Value.ToString(),
                                        WoErpFullNo = table.Cell(i, 3).Value.ToString(),
                                        MtlItemNo = table.Cell(i, 4).Value.ToString(),
                                        Remark = table.Cell(i, 5).Value.ToString(),
                                        QcItemFormatPath = table.Cell(i, 6).Value.ToString(),
                                    };
                                    qcExcelFormats.Add(qcExcelFormat);
                                }
                                catch (Exception ex)
                                {
                                    throw new SystemException("第" + i + "筆資料ERROR:" + ex.Message);
                                }
                            }
                        }
                    }
                    #endregion
                }

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddBatchOutsourcingQcRecord(qcExcelFormats);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    jsonResponse = dataRequestJson;
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

        #region //ImportExcelToAddQcRecordMulti 解析Excel新增QcRecord
        [HttpPost]
        public void ImportExcelToAddQcRecordMulti(string FileId = "")
        {
            try
            {
                if (FileId.Length <= 0) throw new SystemException("必須上傳檔案!!");

                List<QcExcelFormatMulti> qcExcelFormats = new List<QcExcelFormatMulti>();

                string[] FileIds = FileId.Split(',');
                foreach (var fileId in FileIds)
                {
                    #region //取得File資料
                    BmFileInfo fileInfo = qcDataManagementDA.GetFileInfoById(Convert.ToInt32(fileId));
                    if (fileInfo == null)
                    {
                        throw new SystemException("查無此File資料!!");
                    }
                    #endregion

                    #region //解析EXCEL
                    using (var stream = new MemoryStream(fileInfo.FileContent))
                    {
                        using (XLWorkbook workbook = new XLWorkbook(stream))
                        {
                            IXLWorksheet worksheet = workbook.Worksheet(1);
                            var firstCell = worksheet.FirstCellUsed();
                            var lastCell = worksheet.LastCellUsed();

                            int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                            var data = worksheet.Range(firstCell.Address, lastCell.Address);
                            var table = data.AsTable();

                            for (var i = 2; i <= CellLength; i++)
                            {
                                try
                                {
                                    QcExcelFormatMulti qcExcelFormat = new QcExcelFormatMulti()
                                    {
                                        Group = table.Cell(i, 1).Value.ToString(),
                                        QcType = (table.Cell(i, 2).Value.ToString()).Split(':')[0],
                                        WoErpFullNo = table.Cell(i, 3).Value.ToString(),
                                        LotNumber = table.Cell(i, 4).Value.ToString(),
                                        OrderRemark = table.Cell(i, 5).Value.ToString(),
                                        QcItemTemplet = table.Cell(i, 6).Value.ToString(),
                                        QcItemNo = table.Cell(i, 7).Value.ToString(),
                                        QcItemRemark = table.Cell(i, 8).Value.ToString(),
                                        QcItemFormatPath = table.Cell(i, 9).Value.ToString(),
                                        MachineSetting = table.Cell(i, 10).Value.ToString(),
                                    };

                                    // 檢查是否符合條件
                                    if (qcExcelFormat.QcItemTemplet != "" && qcExcelFormat.QcItemNo != "")
                                    {
                                        throw new SystemException("第" + i + "筆資料ERROR: 樣板組合模式和手動添加量測項目模式不可以同時存在");
                                    }
                                    qcExcelFormats.Add(qcExcelFormat);
                                }
                                catch (Exception ex)
                                {
                                    throw new SystemException("第" + i + "筆資料ERROR:" + ex.Message);
                                }
                            }
                        }
                    }
                    #endregion
                }

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.AddBatchQcRecordMulti(qcExcelFormats);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                else
                {
                    jsonResponse = dataRequestJson;
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

        #region //UpdatedOutsourcingQcRecord 更新量測記錄單頭資料
        [HttpPost]
        public void UpdatedOutsourcingQcRecord(int QcRecordId = -1, int QcNoticeId = -1, int QcTypeId = -1,  int MtlItemId = -1, string Remark = "", int CurrentFileId = -1, string QcRecordFile = "", string InputType = "", string SupportAqFlag = "", string ResolveFile = "", string ResolveFileJson = "", string QcRecordFileByNas = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                string ServerPath = Server.MapPath(DataPath);
                string ServerPath2 = Server.MapPath(QcReportDataPath);

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdatedOutsourcingQcRecord(QcRecordId, QcNoticeId, QcTypeId, MtlItemId, Remark, CurrentFileId, QcRecordFile, InputType, ServerPath, ServerPath2, SupportAqFlag, ResolveFile, ResolveFileJson, QcRecordFileByNas);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCheckGoodsOutSourcing 確認收貨狀態
        [HttpPost]
        public void UpdateCheckGoodsOutSourcing(int QcRecordId = -1, string CheckQcMeasureData = "", string Disallowance = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UpdateCheckGoodsOutSourcing(QcRecordId, CheckQcMeasureData, Disallowance);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteQcRecord -- 刪除量測記錄資料
        [HttpPost]
        public void DeleteQcRecord(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "delete");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.DeleteQcRecord(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTempSpreadsheet -- 刪除暫存量測記錄
        [HttpPost]
        public void DeleteTempSpreadsheet(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "delete");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.DeleteTempSpreadsheet(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteResolveFileJson -- 刪除暫存量測解析JSON
        [HttpPost]
        public void DeleteResolveFileJson(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "delete-file");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.DeleteResolveFileJson(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteQcGoodsReceipt -- 刪除進貨檢量測單據
        [HttpPost]
        public void DeleteQcGoodsReceipt(int QcRecordId = -1)
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "delete");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.DeleteQcGoodsReceipt(QcRecordId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteQcRecordFile -- 刪除量測歸檔資料
        [HttpPost]
        public void DeleteQcRecordFile(int QcRecordFileId = -1)
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "delete");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.DeleteQcRecordFile(QcRecordFileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteQcMachineModePlanning -- 刪除量測機台排程資料
        [HttpPost]
        public void DeleteQcMachineModePlanning(int QmmpId = -1)
        {
            try
            {
                //WebApiLoginCheck("QcDataManagement", "delete");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.DeleteQcMachineModePlanning(QmmpId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Upload Data
        #region //UploadQcData 解析SpreadSheet及上傳量測數據
        [HttpPost]
        public void UploadQcData(int QcRecordId = -1, string QcData = "", string BarcodeQcRecordData = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                string ServerPath2 = Server.MapPath(QcReportDataPath);
                dataRequest = qcDataManagementDA.UploadQcData(QcRecordId, QcData, BarcodeQcRecordData, SpreadsheetData, ServerPath2);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UploadTempSpreadsheetJson 暫存SpreadsheetData
        [HttpPost]
        public void UploadTempSpreadsheetJson(int QcRecordId = -1, string UploadJsonString = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UploadTempSpreadsheetJson(QcRecordId, UploadJsonString);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UploadQcGoodsReceipt 上傳進貨檢量測數據
        [HttpPost]
        public void UploadQcGoodsReceipt(int QcRecordId = -1, string QcData = "", string BarcodeQcRecordData = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("QcGoodsReceiptManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                string ServerPath2 = Server.MapPath(QcReportDataPath);
                dataRequest = qcDataManagementDA.UploadQcGoodsReceipt(QcRecordId, QcData, BarcodeQcRecordData, SpreadsheetData, ServerPath2);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UploadQcOutsourcingData 解析SpreadSheet及上傳量測數據
        [HttpPost]
        public void UploadQcOutsourcingData(int QcRecordId = -1, string QcData = "", string BarcodeQcRecordData = "", string SpreadsheetData = "")
        {
            try
            {
                WebApiLoginCheck("QcDataManagement", "update");

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.UploadQcOutsourcingData(QcRecordId, QcData, BarcodeQcRecordData, SpreadsheetData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //ResolveFile 解析量測檔案-主程式
        public void ResolveFile(int QcRecordId = -1, int QcClassId = -1, string ResolveFileJsonString = "", string ResolveType = "")
        {
            if (ResolveFileJsonString.Length <= 0) throw new SystemException("量測檔案資料不能為空!!");

            try
            {
                JObject dataRequestJson = new JObject();
                JObject resolveFileJson = JObject.Parse(ResolveFileJsonString);
                List<ResolveFileOfUA3PModel> resolveFileOfUA3PModels = new List<ResolveFileOfUA3PModel>();
                List<ResolveFileResultModel> resolveFileResults = new List<ResolveFileResultModel>();
                List<QvDesignValueModel> qvDesignValueModels = new List<QvDesignValueModel>();
                List<string> letterNoList = new List<string>();
                int MoId = -1;
                string MtlItemNo = "";
                string QcClassNo = "";
                string UA3PMachineNumber = "";
                string UA3PMachineName = "";
                int UA3PQmmDetailId = -1;
                string UA3PQcEncoding = "UTF-8";
                int CompanyId = -1;

                #region //確認量測單據資料是否正確
                switch (ResolveType)
                {
                    case "IPQC":
                        dataRequest = qcDataManagementDA.GetQcRecord(QcRecordId, "", "", "", "", -1, "", -1, "", "", -1, -1, ""
                            , "", -1, -1);
                        break;
                    case "IQC":
                        dataRequest = qcDataManagementDA.GetQcGoodsReceipt(-1, QcRecordId, "", "", "", "", -1, "", "", "","",-1,"",-1
                            , "", -1, -1);
                        break;
                    default:
                        throw new SystemException("量測類型【" + ResolveType + "】錯誤!!");
                }

                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() == "[]")
                    {
                        throw new SystemException("量測單據資料錯誤!!");
                    }
                    else
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            MoId = Convert.ToInt32(item["MoId"]);
                            MtlItemNo = item["MtlItemNo"].ToString();
                            CompanyId = Convert.ToInt32(item["CompanyId"]);
                        }
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //確認量測類別資料是否正確
                dataRequest = qcDataManagementDA.GetQcClass(QcClassId, "");
                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() == "[]")
                    {
                        throw new SystemException("量測類型資料錯誤!!");
                    }
                    else
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            QcClassNo = item["QcClassNo"].ToString();
                        }
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //取得目前上傳人員資料
                string userNo = "";
                dataRequest = basicInformationDA.GetUser(Convert.ToInt32(HttpContext.Session["UserId"]), -1, -1, "", "", "", "", "", ""
                    , "", -1, -1);

                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            userNo = item["UserNo"].ToString();
                        }
                    }
                    else
                    {
                        throw new SystemException("查詢目前登入人員資料時發生錯誤!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                foreach (var item in resolveFileJson["resolveFileInfo"])
                {
                    #region //確認量測機台資料是否正確
                    if (item["MachineNumber"].ToString().Length <= 0) throw new SystemException("量測機台不能為空!!");
                    dataRequest = qcDataManagementDA.GetQmmDetail(-1, -1, item["MachineNumber"].ToString(), -1, "");
                    dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["data"].ToString() == "[]") throw new SystemException("量測機台資料錯誤!!");

                    int qmmDetailId = -1;
                    string machineName = "";
                    string machineNumber = "";
                    foreach (var item2 in dataRequestJson["data"])
                    {
                        qmmDetailId = Convert.ToInt32(item2["QmmDetailId"]);
                        machineName = item2["MachineName"].ToString();
                        machineNumber = item2["MachineNumber"].ToString();
                    }
                    #endregion

                    #region //若機台為UA3P，個別處理
                    if (machineName.IndexOf("UA3P") != -1)
                    {
                        UA3PQmmDetailId = qmmDetailId;
                        UA3PMachineNumber = item["MachineNumber"].ToString();
                        UA3PMachineName = machineName;

                        ResolveFileOfUA3PModel resolveFileOfUA3PModel = new ResolveFileOfUA3PModel()
                        {
                            FileId = Convert.ToInt32(item["FileId"]),
                            MachineNumber = item["MachineNumber"].ToString(),
                            MachineName = machineName,
                            EffectiveDiameterR1 = item["EffectiveDiameterR1"] != null ? item["EffectiveDiameterR1"].ToString() : "",
                            EffectiveDiameterR2 = item["EffectiveDiameterR2"] != null ? item["EffectiveDiameterR2"].ToString() : "",
                            UserNo = item["UserNo"].ToString()
                        };
                        resolveFileOfUA3PModels.Add(resolveFileOfUA3PModel);
                        continue;
                    }
                    #endregion

                    #region //確認是否有相對應解析方式
                    dataRequest = qcDataManagementDA.GetQcFileParsing(-1, qmmDetailId, "");
                    dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequestJson["data"].ToString() == "[]") throw new SystemException("查無此機台【" + machineName + "】支援的解析方式!!");

                    foreach (var item2 in dataRequestJson["data"])
                    {
                        #region //先確認是否為QV-JSON，若是則額外解析
                        string dataQvJsonRequest = qcDataManagementDA.GetFileInfo(item["FileId"].ToString(), "");
                        JObject dataQvJsonRequestJson = JObject.Parse(dataQvJsonRequest);
                        if (dataQvJsonRequestJson["status"].ToString() == "success")
                        {
                            foreach (var item3 in dataQvJsonRequestJson["data"])
                            {
                                if (item3["FileExtension"].ToString().IndexOf("json") != -1)
                                {
                                    qvDesignValueModels = ResolvesQvDesignValue((byte[])item3["FileContent"]);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            throw new SystemException(dataQvJsonRequestJson["msg"].ToString());
                        }
                        #endregion

                        Type thisType = this.GetType();
                        MethodInfo methodInfo = thisType.GetMethod(item2["FunctionName"].ToString());
                        object[] parameters = new object[] { Convert.ToInt32(item["FileId"]), Convert.ToInt32(item["QcClassId"]), qmmDetailId, item2["Encoding"].ToString(), machineNumber, machineName };
                        List<ResolveFileResultModel> result = (List<ResolveFileResultModel>)methodInfo.Invoke(this, parameters);

                        foreach (var item3 in result)
                        {
                            #region //確認是否有錯誤訊息
                            if (item3.ErrorMessage != null)
                            {
                                throw new SystemException("解析程式【" + item2["FunctionName"].ToString() + "】發生錯誤，訊息如下：" + item3.ErrorMessage);
                            }
                            #endregion

                            #region //此品號、量測項目之是否已經有量測記錄，若有則取其設計值、上下公差
                            double? designValue = -99;
                            double? upperTolerance = -99;
                            double? lowerTolerance = -99;
                            dataRequest = qcDataManagementDA.GetQcMeasureDataHistory(MtlItemNo, item3.QcItemNo);
                            JObject dataRequestQcDataJson = JObject.Parse(dataRequest);
                            if (dataRequestQcDataJson["data"].ToString() != "[]")
                            {
                                foreach (var item4 in dataRequestQcDataJson["data"])
                                {
                                    double checkResult = 0;
                                    designValue = double.TryParse(item4["DesignValue"].ToString(), out checkResult) ? Convert.ToDouble(item4["DesignValue"]) : 0;
                                    upperTolerance = double.TryParse(item4["UpperTolerance"].ToString(), out checkResult) ? Convert.ToDouble(item4["UpperTolerance"]) : 0;
                                    lowerTolerance = double.TryParse(item4["LowerTolerance"].ToString(), out checkResult) ? Convert.ToDouble(item4["LowerTolerance"]) : 0;
                                }
                            }
                            #endregion

                            #region //組合QcItemNo
                            if (item3.MachineNumber == null || item3.MachineNumber == "")
                            {
                                throw new SystemException("解析程式【" + item2["FunctionName"].ToString() + "】尚未維護量測機台編號!!");
                            }
                            string qcItemNo = item3.QcItemNo.Substring(0, 3) + item3.MachineNumber + item3.QcItemNo.ToString().Substring(3, 4);
                            #endregion

                            ResolveFileResultModel resolveFileResult = new ResolveFileResultModel()
                            {
                                QcItemNo = qcItemNo,
                                QcItemName = item3.QcItemName,
                                LetteringNo = item3.LetteringNo,
                                MachineName = item3.MachineName,
                                MeasureValue = item3.MeasureValue,
                                QcItemDesc = item3.QcItemDesc,
                                DesignValue = designValue != -99 ? designValue : item3.DesignValue,
                                UpperTolerance = upperTolerance != -99 ? upperTolerance : item3.UpperTolerance,
                                LowerTolerance = lowerTolerance != -99 ? lowerTolerance : item3.LowerTolerance,
                                UserNo = item3.UserNo != null ? item3.UserNo.ToString() : userNo,
                                ZAxis = item3.ZAxis
                            };

                            resolveFileResults.Add(resolveFileResult);

                            if (letterNoList.IndexOf(item3.LetteringNo) == -1)
                            {
                                letterNoList.Add(item3.LetteringNo);
                            }
                        }
                    }
                    #endregion
                }

                #region //處理UA3P狀況
                if (resolveFileOfUA3PModels.Count > 0)
                {
                    List<ResolveFileResultModel> result = new List<ResolveFileResultModel>();
                    if (CompanyId == 2)
                    {
                        result = ResolvesFileOfUA3P(resolveFileOfUA3PModels, QcClassId, UA3PQmmDetailId, UA3PQcEncoding);
                    }
                    else if (CompanyId == 4)
                    {
                        result = ResolvesFileOfUA3PforJC(resolveFileOfUA3PModels, QcClassId, UA3PQmmDetailId, UA3PQcEncoding);
                    }
                    else
                    {
                        throw new SystemException("此公司別尚未設定UA3P解析程式!!");
                    }

                    foreach (var item3 in result)
                    {
                        #region //確認是否有錯誤訊息
                        if (item3.ErrorMessage != null)
                        {
                            throw new SystemException("解析程式【ResolvesFileOfUA3P】發生錯誤，訊息如下：" + item3.ErrorMessage);
                        }
                        #endregion

                        #region //此品號、量測項目之是否已經有量測記錄，若有則取其設計值、上下公差
                        double? designValue = -99;
                        double? upperTolerance = -99;
                        double? lowerTolerance = -99;
                        dataRequest = qcDataManagementDA.GetQcMeasureDataHistory(MtlItemNo, item3.QcItemNo);
                        JObject dataRequestQcDataJson = JObject.Parse(dataRequest);
                        if (dataRequestQcDataJson["data"].ToString() != "[]")
                        {
                            foreach (var item4 in dataRequestQcDataJson["data"])
                            {
                                designValue = Convert.ToDouble(item4["DesignValue"]);
                                upperTolerance = Convert.ToDouble(item4["UpperTolerance"]);
                                lowerTolerance = Convert.ToDouble(item4["LowerTolerance"]);
                            }
                        }
                        #endregion

                        #region //組合QcItemNo
                        string qcItemNo = item3.QcItemNo.Substring(0, 3) + item3.MachineNumber + item3.QcItemNo.ToString().Substring(3, 4);
                        #endregion

                        ResolveFileResultModel resolveFileResult = new ResolveFileResultModel()
                        {
                            QcItemNo = qcItemNo,
                            QcItemName = item3.QcItemName,
                            LetteringNo = item3.LetteringNo,
                            MachineName = item3.MachineName,
                            MeasureValue = item3.MeasureValue,
                            QcItemDesc = item3.QcItemDesc,
                            DesignValue = designValue != -99 ? designValue : null,
                            UpperTolerance = upperTolerance != -99 ? upperTolerance : null,
                            LowerTolerance = lowerTolerance != -99 ? lowerTolerance : null,
                            UserNo = item3.UserNo
                        };

                        resolveFileResults.Add(resolveFileResult);

                        if (letterNoList.IndexOf(item3.LetteringNo) == -1)
                        {
                            letterNoList.Add(item3.LetteringNo);
                        }
                    }
                }
                #endregion

                #region //處理QV上下限設計值資料
                if (qvDesignValueModels.Count > 0)
                {
                    int count = 0;
                    foreach (var item in resolveFileResults)
                    {
                        foreach (var item2 in qvDesignValueModels)
                        {
                            if (item.QcItemName.IndexOf(item2.QcItemName) != -1)
                            {
                                resolveFileResults[count].DesignValue = item2.DesignValue;
                                resolveFileResults[count].UpperTolerance = item2.UpperTolerance;
                                resolveFileResults[count].LowerTolerance = item2.LowerTolerance;
                                break;
                            }
                        }

                        count++;
                    }
                }
                #endregion

                #region //更新MES.QcRecord ResolveFileJson
                dataRequest = qcDataManagementDA.UpdateResolveFileJson(QcRecordId, ResolveFileJsonString);
                dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException("儲存ResolveFileJson錯誤!!");
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    data = resolveFileResults,
                    letterNo = letterNoList
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

        #region //ResolvesFileOfQV01 解析QV文件-01
        public List<ResolveFileResultModel> ResolvesFileOfQV01(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string MachineNumber = "", string MachineName = "")
        {
            List<ResolveFileResultModel> resolveFileResultModels = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);
                string readTextLine = "";

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            if (item["FileExtension"].ToString().IndexOf("json") != -1)
                            {
                                return resolveFileResultModels;
                            }

                            byte[] fileBytes = (byte[])item["FileContent"];
                            readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //解析QV相關
                #region //開始解析
                string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                int count = 0;
                foreach (var readLine in readTextLineArray)
                {
                    #region //忽略空白
                    if (readLine == "")
                    {
                        count++;
                        continue;
                    }
                    #endregion

                    #region //先抓取量測項目以及刻號
                    string qcItemName = "";
                    string letteringNo = "";
                    double measureValue = 0;
                    if (readLine.IndexOf(":") != -1 && readLine.IndexOf("-") != -1 && readLine.IndexOf("(") != -1)
                    {
                        #region //處理特別格式
                        if (readLine.Split(':')[0].IndexOf("矢高") != -1 && readLine.Split(':')[0].IndexOf("平行度") != -1)
                        {
                            qcItemName = "矢高-平行度-" + readLine.Split(':')[0].Split('-')[0].Split('高')[1];
                            letteringNo = readLine.Split(':')[0].Split('-')[1].Split('-')[0].ToString();
                            if (readLine.Split(':')[1].ToCharArray()[0].ToString() == ".")
                            {
                                measureValue = Convert.ToDouble("0" + readLine.Split(':')[1].ToString());
                            }
                            else
                            {
                                measureValue = Convert.ToDouble(readLine.Split(':')[1].ToString());
                            }

                            GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                        }
                        else if (readLine.Split(':')[0].IndexOf("流道對稱度") != -1)
                        {
                            qcItemName = readLine.Split(':')[0].Split('-')[0];
                            letteringNo = readLine.Split(':')[0].Split('-')[1];

                            if (readLine.Split(':')[1].ToCharArray()[0].ToString() == ".")
                            {
                                measureValue = Convert.ToDouble("0" + readLine.Split(':')[1].ToString());
                            }
                            else
                            {
                                measureValue = Convert.ToDouble(readLine.Split(':')[1].ToString());
                            }

                            GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                        }
                        else
                        {
                            qcItemName = readLine.Split(':')[1].Split('(')[0].Split('-')[0].ToString();
                            letteringNo = readLine.Split(':')[1].Split('(')[0].Split('-')[1].ToString();
                        }
                        #endregion

                        #region //確認後面是否有數值
                        int i = 1;
                        while (readTextLineArray[count + i].IndexOf("=") != -1)
                        {
                            string[] qcContentList = ReplaceEmpty(readTextLineArray[count + i]).Split('=');
                            string qcContent = "";
                            if (qcContentList.Length < 2)
                            {
                                errorQcItemList.Add(ReplaceEmpty(readTextLineArray[count + i]));
                                continue;
                            }
                            else
                            {
                                qcContent = qcContentList[0];
                                measureValue = Convert.ToDouble(qcContentList[1]);
                            }

                            if (qcItemName.IndexOf("口徑") != -1)
                            {
                                if (qcContent.IndexOf("真圓度") != -1)
                                {
                                    string itemNoNumber = GetQcItemNumber(qcItemName, '徑');
                                    GetQcItemPrinciple("口徑真圓度" + itemNoNumber, letteringNo, measureValue);
                                }
                                else if (qcContent.IndexOf("直徑") != -1)
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else if (qcItemName.IndexOf("砂面內徑") != -1)
                            {
                                if (qcContent.IndexOf("直徑") != -1)
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else if (qcItemName.IndexOf("流道寬") != -1)
                            {
                                if (qcContent == "LC")
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else if (qcItemName.IndexOf("EP孔") != -1)
                            {
                                if (qcContent == "X座標")
                                {
                                    string currentQcItemName = "X_" + qcItemName;
                                    GetQcItemPrinciple(ReplaceEmpty(currentQcItemName), letteringNo, measureValue);
                                }
                                else if (qcContent == "Y座標")
                                {
                                    string currentQcItemName = "Y_" + qcItemName;
                                    GetQcItemPrinciple(ReplaceEmpty(currentQcItemName), letteringNo, measureValue);
                                }
                                else if (qcItemName.IndexOf("轉置EP孔") != -1)
                                {
                                    if (qcContent == "直徑")
                                    {
                                        GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                    }
                                    else if (qcContent == "圓心度")
                                    {
                                        string currentQcItemName = "圓度_" + qcItemName;
                                        GetQcItemPrinciple(ReplaceEmpty(currentQcItemName), letteringNo, measureValue);
                                    }
                                }
                            }
                            else
                            {
                                if (qcContent.IndexOf("半徑") != -1 || qcContent.IndexOf("真圓度") != -1 || qcContent.IndexOf("DX") != -1 || qcContent.IndexOf("DY") != -1)
                                {
                                    i++;
                                    continue;
                                }
                                else
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            i++;
                        }
                        #endregion
                    }
                    #endregion
                    count++;
                }
                #endregion

                #region //去空格
                string ReplaceEmpty(string replaceText)
                {
                    replaceText = Regex.Replace(replaceText, @"\s+", String.Empty);
                    return replaceText;
                }
                #endregion

                #region //取得量測項目後數字
                string GetQcItemNumber(string qcItemName, char splitString)
                {
                    string itemNumber = "1";
                    string[] itemNumberArray = qcItemName.Split(splitString);
                    if (itemNumberArray.Length <= 1)
                    {
                        itemNumber = "";
                    }
                    else
                    {
                        itemNumber = itemNumberArray[1];
                    }

                    return itemNumber;
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string principleDesc = "", string letteringNo = "", double measureValue = -1, string relationshipItem = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, "", principleDesc, QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["PrincipleDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MeasureValue = measureValue.ToString(),
                            QcItemDesc = qcItemDesc,
                            MachineNumber = MachineNumber
                        };

                        resolveFileResultModels.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModels.Add(resolveFileResultModel);
                #endregion

                logger.Error(e.Message);
            }

            return resolveFileResultModels;
        }
        #endregion

        #region //ResolvesQvDesignValue 解析QV-JSON
        public List<QvDesignValueModel> ResolvesQvDesignValue(byte[] fileBytes)
        {
            List<QvDesignValueModel> qvDesignValueModels = new List<QvDesignValueModel>();
            Dictionary<string, string> translateDic = new Dictionary<string, string>();

            try
            {
                #region //開啟對照JSON
                string filePath = @"\\192.168.20.199\mes_data\qv-translate.json";
                if (System.IO.File.Exists(filePath))
                {
                    // 讀取JSON檔案的內容
                    string jsonContent = System.IO.File.ReadAllText(filePath);
                    JObject qvTranslateJson = JObject.Parse(jsonContent);
                    foreach (var item in qvTranslateJson)
                    {
                        if (translateDic.ContainsKey(item.Value.ToString()))
                        {
                            continue;
                        }

                        translateDic.Add(item.Value.ToString(), item.Key.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("檔案不存在: " + filePath);
                }
                #endregion

                #region //確認檔案資料是否正確
                string readTextLine = Encoding.GetEncoding("UTF-8").GetString(fileBytes);
                #endregion

                #region //string to json
                JObject qvJson = JObject.Parse(readTextLine);
                #endregion

                #region //resolve json
                foreach (var item in qvJson)
                {
                    string itemNo = item.Key.ToString();
                    if (translateDic.ContainsKey(itemNo))
                    {
                        JObject itemDetailJson = JObject.Parse(item.Value.ToString());

                        QvDesignValueModel qvDesignValueModel = new QvDesignValueModel()
                        {
                            QcItemName = translateDic[itemNo].ToString(),
                            DesignValue = Convert.ToDouble(itemDetailJson["measurementItem"]),
                            UpperTolerance = Convert.ToDouble(itemDetailJson["upperTolerance"]),
                            LowerTolerance = Convert.ToDouble(itemDetailJson["lowerTolerance"])
                        };

                        qvDesignValueModels.Add(qvDesignValueModel);
                    }
                }
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                QvDesignValueModel qvDesignValueModel = new QvDesignValueModel()
                {
                    ErrorMessage = e.Message
                };

                qvDesignValueModels.Add(qvDesignValueModel);
                #endregion

                logger.Error(e.Message);
            }

            return qvDesignValueModels;
        }
        #endregion

        #region //ResolvesFileOfQV01ForJC 解析QV文件-01 FOR 晶彩
        public List<ResolveFileResultModel> ResolvesFileOfQV01ForJC(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string MachineNumber = "", string MachineName = "")
        {
            List<ResolveFileResultModel> resolveFileResultModels = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);
                string readTextLine = "";

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            if (item["FileExtension"].ToString().IndexOf("json") != -1)
                            {
                                return resolveFileResultModels;
                            }

                            byte[] fileBytes = (byte[])item["FileContent"];
                            readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //解析QV相關
                #region //開始解析
                string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                int count = 0;
                foreach (var readLine in readTextLineArray)
                {
                    #region //忽略空白
                    if (readLine == "")
                    {
                        count++;
                        continue;
                    }
                    #endregion

                    #region //先抓取量測項目以及刻號
                    string qcItemName = "";
                    string letteringNo = "";
                    double measureValue = 0;
                    if (readLine.IndexOf(":") != -1 && readLine.IndexOf("-") != -1 && readLine.IndexOf("(") != -1)
                    {
                        #region //處理特別格式
                        if (readLine.Split(':')[0].IndexOf("矢高") != -1 && readLine.Split(':')[0].IndexOf("平行度") != -1)
                        {
                            qcItemName = "矢高-平行度-" + readLine.Split(':')[0].Split('-')[0].Split('高')[1];
                            letteringNo = readLine.Split(':')[0].Split('-')[1].Split('-')[0].ToString();
                            if (readLine.Split(':')[1].ToCharArray()[0].ToString() == ".")
                            {
                                measureValue = Convert.ToDouble("0" + readLine.Split(':')[1].ToString());
                            }
                            else
                            {
                                measureValue = Convert.ToDouble(readLine.Split(':')[1].ToString());
                            }

                            GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                        }
                        else if (readLine.Split(':')[0].IndexOf("流道对称度") != -1)
                        {
                            qcItemName = readLine.Split(':')[0].Split('-')[0];
                            letteringNo = readLine.Split(':')[0].Split('-')[1];

                            if (readLine.Split(':')[1].ToCharArray()[0].ToString() == ".")
                            {
                                measureValue = Convert.ToDouble("0" + readLine.Split(':')[1].ToString());
                            }
                            else
                            {
                                measureValue = Convert.ToDouble(readLine.Split(':')[1].ToString());
                            }

                            GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                        }
                        else
                        {
                            qcItemName = readLine.Split(':')[1].Split('(')[0].Split('-')[0].ToString();
                            letteringNo = readLine.Split(':')[1].Split('(')[0].Split('-')[1].ToString();
                        }
                        #endregion

                        #region //確認後面是否有數值
                        int i = 1;
                        while (readTextLineArray[count + i].IndexOf("=") != -1)
                        {
                            string[] qcContentList = ReplaceEmpty(readTextLineArray[count + i]).Split('=');
                            string qcContent = "";
                            if (qcContentList.Length < 2)
                            {
                                errorQcItemList.Add(ReplaceEmpty(readTextLineArray[count + i]));
                                continue;
                            }
                            else
                            {
                                qcContent = qcContentList[0];
                                measureValue = Convert.ToDouble(qcContentList[1]);
                            }

                            if (qcItemName.IndexOf("口径") != -1)
                            {
                                if (qcContent.IndexOf("圆度") != -1)
                                {
                                    string itemNoNumber = GetQcItemNumber(qcItemName, "径");
                                    if (itemNoNumber == "") itemNoNumber = "1";
                                    GetQcItemPrinciple("口径真圆度" + itemNoNumber, letteringNo, measureValue);
                                }
                                else if (qcContent.IndexOf("直径") != -1)
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else if (qcItemName.IndexOf("砂面内径") != -1)
                            {
                                if (qcContent.IndexOf("直径") != -1)
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else if (qcItemName.IndexOf("流道宽") != -1)
                            {
                                if (qcContent == "LC")
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else if (qcItemName.IndexOf("EP孔") != -1)
                            {
                                if (qcContent == "X座标")
                                {
                                    string currentQcItemName = "X_" + qcItemName;
                                    GetQcItemPrinciple(ReplaceEmpty(currentQcItemName), letteringNo, measureValue);
                                }
                                else if (qcContent == "Y座标")
                                {
                                    string currentQcItemName = "Y_" + qcItemName;
                                    GetQcItemPrinciple(ReplaceEmpty(currentQcItemName), letteringNo, measureValue);
                                }
                                else if (qcItemName.IndexOf("转置EP孔") != -1)
                                {
                                    if (qcContent == "直径")
                                    {
                                        GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                    }
                                    else if (qcContent == "圆心度")
                                    {
                                        string currentQcItemName = "圆度_" + qcItemName;
                                        GetQcItemPrinciple(ReplaceEmpty(currentQcItemName), letteringNo, measureValue);
                                    }
                                }
                            }
                            else if (qcItemName.IndexOf("精端面平面度") != -1)
                            {
                                if (qcContent == "平面度")
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            else
                            {
                                if (qcContent.IndexOf("半径") != -1 || qcContent.IndexOf("真圆度") != -1 || qcContent.IndexOf("DX") != -1 || qcContent.IndexOf("DY") != -1)
                                {
                                    i++;
                                    continue;
                                }
                                else
                                {
                                    GetQcItemPrinciple(ReplaceEmpty(qcItemName), letteringNo, measureValue);
                                }
                            }
                            i++;
                        }
                        #endregion
                    }
                    #endregion
                    count++;
                }
                #endregion

                #region //去空格
                string ReplaceEmpty(string replaceText)
                {
                    replaceText = Regex.Replace(replaceText, @"\s+", String.Empty);
                    return replaceText;
                }
                #endregion

                #region //取得量測項目後數字
                string GetQcItemNumber(string qcItemName, string splitString)
                {
                    char ch = splitString[0];
                    string itemNumber = "1";
                    string[] itemNumberArray = qcItemName.Split(ch);
                    if (itemNumberArray.Length <= 1)
                    {
                        itemNumber = "";
                    }
                    else
                    {
                        itemNumber = itemNumberArray[1];
                    }

                    return itemNumber;
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string principleDesc = "", string letteringNo = "", double measureValue = -1, string relationshipItem = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, "", principleDesc, QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["PrincipleDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MeasureValue = measureValue.ToString(),
                            QcItemDesc = qcItemDesc,
                            MachineNumber = MachineNumber
                        };

                        resolveFileResultModels.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModels.Add(resolveFileResultModel);
                #endregion

                logger.Error(e.Message);
            }

            return resolveFileResultModels;
        }
        #endregion

        #region //ResolvesFileOfUA3P 解析UA3P文件 for 中揚
        public List<ResolveFileResultModel> ResolvesFileOfUA3P(List<ResolveFileOfUA3PModel> inputQcDataJson, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "")
        {
            List<ResolveFileResultModel> resolveFileResultModels = new List<ResolveFileResultModel>();
            List<JObject> qcDataList = new List<JObject>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            try
            {
                #region //解開inputQcDataJson JSON資料包
                List<string> FileId = new List<string>();
                List<string> MachineNumber = new List<string>();
                List<string> MachineName = new List<string>();
                List<string> EffectiveDiameterR1 = new List<string>();
                List<string> EffectiveDiameterR2 = new List<string>();
                List<string> UserNo = new List<string>();

                foreach (var item in inputQcDataJson)
                {
                    FileId.Add(item.FileId.ToString());
                    MachineNumber.Add(item.MachineNumber.ToString());
                    MachineName.Add(item.MachineName.ToString());
                    EffectiveDiameterR1.Add(item.EffectiveDiameterR1.ToString());
                    EffectiveDiameterR2.Add(item.EffectiveDiameterR2.ToString());
                    UserNo.Add(item.UserNo.ToString());

                }
                #endregion

                #region //確認檔案資料是否正確, 並取得所有的檔案名稱
                //List<string> FileIdList = new List<string>();
                //FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileId.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);
                short fileCount = 0;

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var itemA in dataRequestJson["data"])
                        {
                            #region 解析UA3P相關
                            string readTextLineA = "";
                            string caveNo = "";
                            short laCount = 0;
                            string[] measureValue;
                            string serialCode = "";

                            //string filePureName = itemA["FileName"].ToString();

                            string filenameA = itemA["FileFullName"].ToString();
                            caveNo = filenameA.Split('-')[2].Split('.')[0];
                            string fileName = filenameA;
                                                        
                            if (fileName.IndexOf("decenter") != -1)
                            {
                                fileName = fileName.Replace("-decenter", "");
                            }
                            else if (fileName.IndexOf("ra") != -1)
                            {
                                fileName = fileName.Replace("-ra", "");
                            }

                            if (fileName != "H.txt")
                            {
                                string[] spName = fileName.Split('-');
                                if (spName.Length > 3)
                                {
                                    //檔名切割，要注意decenter、ra
                                    caveNo = spName[2];
                                    laCount = short.Parse(spName[3].Split('.')[0]);
                                }
                                else
                                {
                                    caveNo = spName[2].Split('.')[0];
                                    laCount = 1;
                                }
                            }

                            byte[] fileBytesA = (byte[])itemA["FileContent"];
                            readTextLineA = Encoding.GetEncoding(QcEncoding).GetString(fileBytesA);
                            //string a = MachineNumber[fileCount];
                            //PV、AS
                            if (filenameA.IndexOf("R2") != -1)
                            {
                                laCount += 1;
                            }
                            if (filenameA.IndexOf(".alx.txt") != -1 || filenameA.IndexOf(".aly.txt") != -1)
                            {                                
                                #region 取得 PV： .alx.txt
                                if (filenameA.IndexOf(".alx.txt") != -1)
                                {
                                    measureValue = GetUa3pValueTxt(readTextLineA);
                                    serialCode = "l0" + string.Format("{0:00}", laCount); //PV's Code
                                    GetQcItemPrinciple(serialCode, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                }
                                #endregion

                                #region 精車模仁要多抓一個 y 軸的 PV： .aly.txt
                                if (QcClassId == 2)
                                {
                                    foreach (var itemB in dataRequestJson["data"])
                                    {
                                        string filenameB = itemA["FileFullName"].ToString();
                                        if (filenameB == filenameA.Replace("alx.txt", "aly.txt"))
                                        {
                                            byte[] fileBytesB = (byte[])itemA["FileContent"];
                                            string readTextLineB = Encoding.GetEncoding("UTF-8").GetString(fileBytesB);
                                            measureValue = GetUa3pValueTxt(readTextLineB);
                                            serialCode = "l002"; //PV 
                                            GetQcItemPrinciple(serialCode, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                            readTextLineB = "";
                                        }
                                    }
                                }
                                #endregion

                                #region 若有效徑 effectivePathR1、effectivePathR1 不為空，則計算AS
                                if (EffectiveDiameterR1[fileCount] != "" && EffectiveDiameterR2[fileCount] != "")
                                {
                                    string effectivePath = "";
                                    string serialCodeAS = "";

                                    #region 設定有效徑
                                    if (filenameA.IndexOf("R1") != -1)
                                    {
                                        effectivePath = EffectiveDiameterR1[fileCount];
                                    }
                                    else if (filenameA.IndexOf("R2") != -1)
                                    {
                                        effectivePath = EffectiveDiameterR2[fileCount];
                                    }
                                    #endregion

                                    foreach (var itemB in dataRequestJson["data"])
                                    {
                                        string filenameB = itemA["FileFullName"].ToString();
                                        if (filenameB == filenameA.Replace("alx", "aly"))
                                        {
                                            byte[] fileBytesB = (byte[])itemA["FileContent"];

                                            string readTextLineB = Encoding.GetEncoding("UTF-8").GetString(fileBytesB);
                                            measureValue = Ua3pCalculateAS(readTextLineA, readTextLineB, effectivePath, serialCodeAS);
                                            serialCodeAS = "l1" + string.Format("{0:00}", laCount); //AS
                                            GetQcItemPrinciple(serialCodeAS, caveNo, measureValue[0], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                            readTextLineB = "";
                                        }
                                    }
                                }
                                #endregion

                            }
                            //BestFit(ACC,PV)、BestFitR
                            else if (filenameA.IndexOf(".alx.BF.txt") != -1)
                            {
                                #region 取得 BestFit： .alx.BF.txt
                                measureValue = GetUa3pValueTxt(readTextLineA);
                                serialCode = "l4" + string.Format("{0:00}", laCount); //ACC(PV、BestFit)
                                GetQcItemPrinciple(serialCode, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);

                                string serialCodeBestFitR = "k2" + string.Format("{0:00}", laCount); //BestFitR
                                GetQcItemPrinciple(serialCodeBestFitR, caveNo, measureValue[3], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                #endregion
                            }
                            //UA3P RA
                            else if (filenameA.IndexOf("ra.csv") != -1)
                            {
                                #region 取得 Ra： ra.csv
                                measureValue = GetUa3pValueCsv(readTextLineA);
                                string serialCodeRa = "l2" + string.Format("{0:00}", laCount); //RA
                                GetQcItemPrinciple(serialCodeRa, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                #endregion
                            }
                            //UA3P Decenter
                            else if (filenameA.IndexOf("decenter.csv") != -1)
                            {
                                #region 取得 Decenter：decenter.csv
                                measureValue = GetUa3pValueCsv(readTextLineA);
                                string serialCodeDecenter = "k1" + string.Format("{0:00}", laCount); //Decenter measureValue[2]
                                GetQcItemPrinciple(serialCodeDecenter, caveNo, measureValue[0], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);

                                string serialCodeTilt = "k0" + string.Format("{0:00}", laCount); //Tilt measureValue[3]
                                GetQcItemPrinciple(serialCodeTilt, caveNo, measureValue[1], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                #endregion
                            }
                            //UA3P 矢高
                            else if (filenameA == "H.txt")
                            {
                                #region 取得矢高：H.txt
                                string[] SagittalValue = GetUa3pHtxt(readTextLineA);
                                #endregion
                            }
                            #endregion
                            readTextLineA = "";
                            fileCount++;
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string serialCode = "", string letteringNo = "", string measureValue = "", string Machinenumber = "", string Machinename = "", string Userno = null)
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["QcItemDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MeasureValue = measureValue,
                            QcItemDesc = qcItemDesc,
                            MachineNumber = Machinenumber,
                            MachineName = Machinename,
                            UserNo = Userno
                        };

                        resolveFileResultModels.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion

                #region 解析 .txt 文件: PV ACC BestFitR
                string[] GetUa3pValueTxt(string readTextLine)
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                    string ua3pMachine = "";
                    string effectivePath = "";
                    string measureBestFitR = "";
                    string measureValuePV = "";
                    string[] returnMeasureValue = new string[4];

                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        #region //.txt,.BF.txt 檔案 包含 PV,AS
                        if (readLine.IndexOf(":") != -1)
                        {
                            string[] splitReadLine = readLine.Split(':');
                            #region //取得值
                            if (splitReadLine[0].IndexOf("Machine number") != -1)
                            {
                                //ua3p 機台編號
                                ua3pMachine = splitReadLine[1].Trim().ToString();
                            }
                            else if (splitReadLine[0].IndexOf("NC Type") != -1)
                            {
                                //量測有效徑
                                if (splitReadLine[1].IndexOf("X Linear") != -1)
                                {
                                    effectivePath = splitReadLine[1].Split('-')[1].Replace(')', ' ').Trim().ToString();
                                }
                            }
                            else if (splitReadLine[0].IndexOf("Alignment type") != -1)
                            {
                                //bestfitR
                                if (splitReadLine[1].IndexOf("BestFitR") != -1)
                                {
                                    measureBestFitR = splitReadLine[1].Split('=')[1].Split('(')[0].Trim().ToString();
                                }
                            }
                            else if (splitReadLine[0].IndexOf("P-V") != -1)
                            {
                                //PV,FIG,ACC
                                measureValuePV = splitReadLine[1].Split('(')[0].Trim().ToString();
                            }
                            #endregion
                        }

                        #endregion
                        count++;
                    }
                    returnMeasureValue[0] = ua3pMachine; //ua3p 機台編號
                    returnMeasureValue[1] = effectivePath; //量測有效徑
                    returnMeasureValue[2] = measureValuePV; //PV,FIG,ACC
                    returnMeasureValue[3] = measureBestFitR; //BestFitR、Tilt
                    return returnMeasureValue;

                }
                #endregion

                #region 解析 AS，文件：.alx.txt,.aly.txt
                string[] Ua3pCalculateAS(string readTextLineA, string readTextLineB, string effectivePath = "", string serialCodeAS = "")
                {
                    string readTextLine = "";
                    string[] returnMeasureValue = new string[4];
                    string[] arrayPointX = new string[0] { };
                    string[] arrayPointY = new string[0] { };
                    int xCount = 0;
                    int yCount = 0;
                    double CalcValue = 0;
                    double maxmValue = 0;
                    double miniValue = 1000;
                    double averageValue = 0;

                    #region 取得 .alx、.aly 的點資料，放入陣列 arrayPointX、arrayPointY
                    for (int i = 0; i < 2; i++)
                    {
                        readTextLine = new[] { readTextLineA, readTextLineB }[i];
                        string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                        int count = 0;
                        foreach (var readLine in readTextLineArray)
                        {
                            #region //忽略空白
                            if (readLine == "")
                            {
                                count++;
                                continue;
                            }
                            #endregion

                            if (readLine.IndexOf(":") != -1)
                            {
                                string[] splitReadLine = readLine.Split(':');
                                if (splitReadLine[0].IndexOf("NC Type") != -1)
                                {
                                    //取得量測有效徑
                                    if (splitReadLine[1].IndexOf("X Linear") != -1)
                                    {
                                        //effectivePath = (float.Parse(splitReadLine[1].Split('-')[1].Replace(')', ' ').Trim()) * 0.95).ToString();
                                    }
                                }
                                if (splitReadLine[1].StartsWith("("))
                                {
                                    //點資料：1 :(    0.2839058116,   -0.0000035886,   -0.0314491796,   -0.0000694893 )
                                    //括號內的數字，分別為：X，Y，Z，offset
                                    string[] splitPointValue = splitReadLine[1].Replace('(', ' ').Replace(')', ' ').Trim().Split(',');

                                    // i=1：.alx.txt，i=2：.aly.txt
                                    if (i == 0)
                                    {
                                        //X 點資料，X小於等於有效徑，才將offset的值放進陣列
                                        if (Math.Abs(double.Parse(splitPointValue[0])) <= double.Parse(effectivePath))
                                        {
                                            Array.Resize(ref arrayPointX, arrayPointX.Length + 1);
                                            arrayPointX[xCount] = splitPointValue[3];
                                            xCount++;
                                        }
                                    }
                                    else if (i == 1)
                                    {
                                        //Y 點資料，Y小於等於有效徑，才將offset的值放進陣列
                                        if (Math.Abs(double.Parse(splitPointValue[1])) <= double.Parse(effectivePath))
                                        {
                                            Array.Resize(ref arrayPointY, arrayPointY.Length + 1);
                                            arrayPointY[yCount] = splitPointValue[3];
                                            yCount++;
                                        }
                                    }
                                }
                            }
                            count++;
                        }
                    }
                    #endregion

                    #region 從陣列arrayPointX、arrayPointY 取得最大值、最小值、平均值
                    if (QcClassId == 7)
                    {
                        //白物鏡片，只取邊境 x-y (arrayPointX[0] - arrayPointY[0]、arrayPointX[xCount - 1] - arrayPointY[yCount - 1]) 的最大值
                        double firstValue = 0;
                        double finalValue = 0;

                        firstValue = Math.Abs(double.Parse(arrayPointX[0]) - double.Parse(arrayPointY[0])) * 1000;
                        finalValue = Math.Abs(double.Parse(arrayPointX[xCount - 1]) - double.Parse(arrayPointY[yCount - 1])) * 1000;

                        if (firstValue > finalValue)
                            maxmValue = firstValue;
                        else
                            maxmValue = finalValue;
                    }
                    else if (QcClassId == 8)
                    {
                        //模造鏡片，取 x-y 最大值
                        int endPoint = 0;
                        int CenterOffset = 0;
                        double sumValue = 0;

                        if (xCount >= yCount) //取最少的點資料
                            endPoint = yCount;
                        else
                            endPoint = xCount;

                        CenterOffset = Math.Abs((xCount - yCount) / 2); //計算xy中心點偏差值，要除以2取整數，無條件捨去，

                        for (int i = 0; i < endPoint; i++)
                        {
                            if (xCount >= yCount)
                                CalcValue = Math.Abs(double.Parse(arrayPointX[i + CenterOffset]) - double.Parse(arrayPointY[i]) * 1000);
                            else
                                CalcValue = Math.Abs(double.Parse(arrayPointX[i]) - double.Parse(arrayPointY[i + CenterOffset]) * 1000);

                            sumValue = sumValue + CalcValue;

                            if (CalcValue > maxmValue) //最大值
                                maxmValue = CalcValue;
                            if (CalcValue < miniValue) //最小值
                                miniValue = CalcValue;
                        }
                        averageValue = sumValue / endPoint; //平均值                        
                    }
                    #endregion
                    returnMeasureValue[0] = maxmValue.ToString();
                    returnMeasureValue[1] = miniValue.ToString();
                    returnMeasureValue[2] = averageValue.ToString();

                    return returnMeasureValue;
                }
                #endregion

                #region 解析 .csv 文件：RA Decenter
                string[] GetUa3pValueCsv(string readTextLine)
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                    string measureValueTilt = "";
                    string measureValueDecenter = "";
                    string measureValueRa = "";
                    string[] returnMeasureValue = new string[3];

                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        #region //.csv 檔案 包含 RA,DECENTER
                        if (readLine.IndexOf(",") != -1)
                        {
                            string[] splitReadLine = readLine.Split(',');
                            if (splitReadLine[0].IndexOf("Ra") != -1)
                            {
                                measureValueRa = splitReadLine[1];
                            }
                            else if (splitReadLine[0].IndexOf("Decenter") != -1)
                            {
                                measureValueDecenter = splitReadLine[1];
                            }
                            else if (splitReadLine[0].IndexOf("Tilt") != -1)
                            {
                                measureValueTilt = (double.Parse(splitReadLine[1]) * 60 + double.Parse(splitReadLine[3])).ToString();
                            }
                        }
                        #endregion
                        count++;
                    }
                    returnMeasureValue[0] = measureValueDecenter; //Decenter
                    returnMeasureValue[1] = measureValueTilt; //Tilt
                    returnMeasureValue[2] = measureValueRa; //Tilt
                    return returnMeasureValue;
                }
                #endregion

                #region 解析矢高 H.txt
                string[] GetUa3pHtxt(string readTextLine)
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                    short arrayCount = 0;
                    string[] returnMeasureValue = new string[4];

                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            arrayCount++;
                            continue;
                        }
                        #endregion

                        if (readLine.IndexOf("-") != -1)
                        {
                            float measureValue = 0;
                            string serialCodeSagittal = "";
                            string caveNoSagittal = "";
                            short laSagittal = 0;
                            caveNoSagittal = readLine.Split('-')[2];
                            //Console.Write("ALL No：" + readLine + ", caveNo：" + caveNoSagittal + "\n");

                            if (readLine.Split('-').Length > 3)
                                laSagittal = short.Parse(readLine.Split('-')[3]);
                            else
                                laSagittal = 1;

                            //判斷同一穴號，是否有第二個矢高
                            short countSagittal = 0;
                            short whileCount = 1;
                            while (readTextLineArray[arrayCount + whileCount].IndexOf("-") == -1)
                            {
                                if (readTextLineArray[arrayCount + whileCount] != "")
                                {
                                    if (readTextLineArray[arrayCount + whileCount].IndexOf("\t") != -1)
                                    {
                                        measureValue = (float.Parse(readTextLineArray[arrayCount + whileCount].Split('\t')[0]) + float.Parse(readTextLineArray[arrayCount + whileCount].Split('\t')[1])) / 2;
                                    }
                                    else
                                    {
                                        measureValue = float.Parse(readTextLineArray[arrayCount + whileCount]);
                                    }

                                    if (readLine.IndexOf("R1") != -1)
                                        serialCodeSagittal = "a1" + string.Format("{0:00}", (laSagittal - 1) * 3 + 1 + countSagittal);
                                    else if (readLine.IndexOf("R2") != -1)
                                        serialCodeSagittal = "a1" + string.Format("{0:00}", (laSagittal - 1) * 3 + 28 + countSagittal);
                                    countSagittal++;

                                    GetQcItemPrinciple(serialCodeSagittal, caveNoSagittal, measureValue.ToString(), MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);

                                    whileCount++;
                                }
                                else
                                {
                                    whileCount++;
                                    continue;
                                }
                            }
                        }
                        arrayCount++;
                    }
                    return returnMeasureValue;
                }
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModels.Add(resolveFileResultModel);
                #endregion

                logger.Error(e.Message);
            }

            return resolveFileResultModels;
        }
        #endregion

        #region //ResolvesFileOfUA3PforJC 解析UA3P文件 for 晶彩
        public List<ResolveFileResultModel> ResolvesFileOfUA3PforJC(List<ResolveFileOfUA3PModel> inputQcDataJson, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "")
        {
            List<ResolveFileResultModel> resolveFileResultModels = new List<ResolveFileResultModel>();
            List<JObject> qcDataList = new List<JObject>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            try
            {
                #region //解開inputQcDataJson JSON資料包
                List<string> FileId = new List<string>();
                List<string> MachineNumber = new List<string>();
                List<string> MachineName = new List<string>();
                List<string> EffectiveDiameterR1 = new List<string>();
                List<string> EffectiveDiameterR2 = new List<string>();
                List<string> UserNo = new List<string>();

                foreach (var item in inputQcDataJson)
                {
                    FileId.Add(item.FileId.ToString());
                    MachineNumber.Add(item.MachineNumber.ToString());
                    MachineName.Add(item.MachineName.ToString());
                    EffectiveDiameterR1.Add(item.EffectiveDiameterR1.ToString());
                    EffectiveDiameterR2.Add(item.EffectiveDiameterR2.ToString());
                    UserNo.Add(item.UserNo.ToString());

                }
                #endregion

                #region //確認檔案資料是否正確, 並取得所有的檔案名稱
                //List<string> FileIdList = new List<string>();
                //FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileId.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);
                short fileCount = 0;

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var itemA in dataRequestJson["data"])
                        {
                            #region 解析UA3P相關
                            string readTextLineA = "";
                            string caveNo = "";
                            short laCount = 0;
                            string[] measureValue;
                            string serialCode = "";

                            //string filePureName = itemA["FileName"].ToString();

                            string filenameA = itemA["FileFullName"].ToString();
                            caveNo = filenameA.Split('-')[1].Split('.')[0];
                            string fileName = filenameA;
                            if (fileName.IndexOf("decenter") != -1)
                            {
                                fileName = fileName.Replace("-decenter", "");
                            }
                            else if (fileName.IndexOf("ra") != -1)
                            {
                                fileName = fileName.Replace("-ra", "");
                            }

                            if (fileName.IndexOf("sag") == -1) //不是矢高的檔案就繼續執行
                            {
                                string[] spName = fileName.Split('-');
                                if (spName.Length > 2)
                                {
                                    //檔名切割，要注意decenter、ra
                                    caveNo = spName[1];
                                    laCount = short.Parse(spName[2].Split('.')[0]);
                                }
                                else
                                {
                                    caveNo = spName[1].Split('.')[0];
                                    laCount = 1;
                                }
                            }

                            byte[] fileBytesA = (byte[])itemA["FileContent"];
                            readTextLineA = Encoding.GetEncoding(QcEncoding).GetString(fileBytesA);
                            //string a = MachineNumber[fileCount];
                            //PV、AS
                            if (filenameA.IndexOf(".alx.txt") != -1 || filenameA.IndexOf(".aly.txt") != -1)
                            {
                                if (filenameA.IndexOf("R2") != -1)
                                {
                                    laCount += 1;
                                }
                                #region 取得 PV： .alx.txt
                                if (filenameA.IndexOf(".alx.txt") != -1)
                                {
                                    measureValue = GetUa3pValueTxt(readTextLineA);
                                    serialCode = "l0" + string.Format("{0:00}", laCount); //PV's Code
                                    GetQcItemPrinciple(serialCode, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                }
                                #endregion

                                #region 精車模仁要多抓一個 y 軸的 PV： .aly.txt
                                if (QcClassId == 39)
                                {
                                    foreach (var itemB in dataRequestJson["data"])
                                    {
                                        string filenameB = itemA["FileFullName"].ToString();
                                        if (filenameB == filenameA.Replace("alx.txt", "aly.txt"))
                                        {
                                            byte[] fileBytesB = (byte[])itemA["FileContent"];
                                            string readTextLineB = Encoding.GetEncoding("UTF-8").GetString(fileBytesB);
                                            measureValue = GetUa3pValueTxt(readTextLineB);
                                            serialCode = "l002"; //PV 
                                            GetQcItemPrinciple(serialCode, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                            readTextLineB = "";
                                        }
                                    }
                                }
                                #endregion

                                #region 若有效徑 effectivePathR1、effectivePathR1 不為空，則計算AS
                                if (EffectiveDiameterR1[fileCount] != "" && EffectiveDiameterR2[fileCount] != "")
                                {
                                    string effectivePath = "";
                                    string serialCodeAS = "";

                                    #region 設定有效徑
                                    if (filenameA.IndexOf("R1") != -1)
                                    {
                                        effectivePath = EffectiveDiameterR1[fileCount];
                                    }
                                    else if (filenameA.IndexOf("R2") != -1)
                                    {
                                        effectivePath = EffectiveDiameterR2[fileCount];
                                    }
                                    #endregion

                                    foreach (var itemB in dataRequestJson["data"])
                                    {
                                        string filenameB = itemA["FileFullName"].ToString();
                                        if (filenameB == filenameA.Replace("alx", "aly"))
                                        {
                                            byte[] fileBytesB = (byte[])itemA["FileContent"];

                                            string readTextLineB = Encoding.GetEncoding("UTF-8").GetString(fileBytesB);
                                            measureValue = Ua3pCalculateAS(readTextLineA, readTextLineB, effectivePath, serialCodeAS);
                                            serialCodeAS = "l1" + string.Format("{0:00}", laCount); //AS
                                            GetQcItemPrinciple(serialCodeAS, caveNo, measureValue[0], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                            readTextLineB = "";
                                        }
                                    }
                                }
                                #endregion

                            }
                            //BestFit(ACC,PV)、BestFitR
                            else if (filenameA.IndexOf(".alx.BF.txt") != -1)
                            {
                                #region 取得 BestFit： .alx.BF.txt
                                measureValue = GetUa3pValueTxt(readTextLineA);
                                serialCode = "l4" + string.Format("{0:00}", laCount); //ACC(PV、BestFit)
                                GetQcItemPrinciple(serialCode, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);

                                string serialCodeBestFitR = "k2" + string.Format("{0:00}", laCount); //BestFitR
                                GetQcItemPrinciple(serialCodeBestFitR, caveNo, measureValue[3], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                #endregion
                            }
                            //UA3P RA
                            else if (filenameA.IndexOf("ra.csv") != -1)
                            {
                                #region 取得 Ra： ra.csv
                                measureValue = GetUa3pValueCsv(readTextLineA);
                                string serialCodeRa = "l2" + string.Format("{0:00}", laCount); //RA
                                GetQcItemPrinciple(serialCodeRa, caveNo, measureValue[2], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                #endregion
                            }
                            //UA3P Decenter
                            else if (filenameA.IndexOf("decenter.csv") != -1)
                            {
                                #region 取得 Decenter：decenter.csv
                                measureValue = GetUa3pValueCsv(readTextLineA);
                                string serialCodeDecenter = "k1" + string.Format("{0:00}", laCount); //Decenter measureValue[2]
                                GetQcItemPrinciple(serialCodeDecenter, caveNo, measureValue[0], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);

                                string serialCodeTilt = "k0" + string.Format("{0:00}", laCount); //Tilt measureValue[3]
                                GetQcItemPrinciple(serialCodeTilt, caveNo, measureValue[1], MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                #endregion
                            }
                            //UA3P 矢高
                            else if (filenameA.IndexOf("sag.xlsx") != -1)
                            {
                                #region 取得矢高：sag.xlsx
                                string filePath = @"\\192.168.20.199\d8-系統開發室\99-單位共用區\滕峮\檔案測試\ua3pHeightTempforJC.xlsx";
                                byte[] fileBytes = (byte[])itemA["FileContent"];
                                string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                                ////將byte[] 儲存至預設路徑 filePath
                                System.IO.File.WriteAllBytes(filePath, fileBytes);
                                //資料解析
                                GetExcel(filePath);
                                #endregion
                            }
                            #endregion
                            readTextLineA = "";
                            fileCount++;
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string serialCode = "", string letteringNo = "", string measureValue = "", string Machinenumber = "", string Machinename = "", string Userno = null)
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["QcItemDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MeasureValue = measureValue,
                            QcItemDesc = qcItemDesc,
                            MachineNumber = Machinenumber,
                            MachineName = Machinename,
                            UserNo = Userno
                        };

                        resolveFileResultModels.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion

                #region 解析 .txt 文件: PV ACC BestFitR
                string[] GetUa3pValueTxt(string readTextLine)
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                    string ua3pMachine = "";
                    string effectivePath = "";
                    string measureBestFitR = "";
                    string measureValuePV = "";
                    string[] returnMeasureValue = new string[4];

                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        #region //.txt,.BF.txt 檔案 包含 PV,AS
                        if (readLine.IndexOf(":") != -1)
                        {
                            string[] splitReadLine = readLine.Split(':');
                            #region //取得值
                            if (splitReadLine[0].IndexOf("Machine number") != -1)
                            {
                                //ua3p 機台編號
                                ua3pMachine = splitReadLine[1].Trim().ToString();
                            }
                            else if (splitReadLine[0].IndexOf("NC Type") != -1)
                            {
                                //量測有效徑
                                if (splitReadLine[1].IndexOf("X Linear") != -1)
                                {
                                    effectivePath = splitReadLine[1].Split('-')[1].Replace(')', ' ').Trim().ToString();
                                }
                            }
                            else if (splitReadLine[0].IndexOf("Alignment type") != -1)
                            {
                                //bestfitR
                                if (splitReadLine[1].IndexOf("BestFitR") != -1)
                                {
                                    measureBestFitR = splitReadLine[1].Split('=')[1].Split('(')[0].Trim().ToString();
                                }
                            }
                            else if (splitReadLine[0].IndexOf("P-V") != -1)
                            {
                                //PV,FIG,ACC
                                measureValuePV = splitReadLine[1].Split('(')[0].Trim().ToString();
                            }
                            #endregion
                        }

                        #endregion
                        count++;
                    }
                    returnMeasureValue[0] = ua3pMachine; //ua3p 機台編號
                    returnMeasureValue[1] = effectivePath; //量測有效徑
                    returnMeasureValue[2] = measureValuePV; //PV,FIG,ACC
                    returnMeasureValue[3] = measureBestFitR; //BestFitR、Tilt
                    return returnMeasureValue;

                }
                #endregion

                #region 解析 AS，文件：.alx.txt,.aly.txt
                string[] Ua3pCalculateAS(string readTextLineA, string readTextLineB, string effectivePath = "", string serialCodeAS = "")
                {
                    string readTextLine = "";
                    string[] returnMeasureValue = new string[4];
                    string[] arrayPointX = new string[0] { };
                    string[] arrayPointY = new string[0] { };
                    int xCount = 0;
                    int yCount = 0;
                    double CalcValue = 0;
                    double maxmValue = 0;
                    double miniValue = 1000;
                    double averageValue = 0;

                    #region 取得 .alx、.aly 的點資料，放入陣列 arrayPointX、arrayPointY
                    for (int i = 0; i < 2; i++)
                    {
                        readTextLine = new[] { readTextLineA, readTextLineB }[i];
                        string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                        int count = 0;
                        foreach (var readLine in readTextLineArray)
                        {
                            #region //忽略空白
                            if (readLine == "")
                            {
                                count++;
                                continue;
                            }
                            #endregion

                            if (readLine.IndexOf(":") != -1)
                            {
                                string[] splitReadLine = readLine.Split(':');
                                if (splitReadLine[0].IndexOf("NC Type") != -1)
                                {
                                    //取得量測有效徑
                                    if (splitReadLine[1].IndexOf("X Linear") != -1)
                                    {
                                        //effectivePath = (float.Parse(splitReadLine[1].Split('-')[1].Replace(')', ' ').Trim()) * 0.95).ToString();
                                    }
                                }
                                if (splitReadLine[1].StartsWith("("))
                                {
                                    //點資料：1 :(    0.2839058116,   -0.0000035886,   -0.0314491796,   -0.0000694893 )
                                    //括號內的數字，分別為：X，Y，Z，offset
                                    string[] splitPointValue = splitReadLine[1].Replace('(', ' ').Replace(')', ' ').Trim().Split(',');

                                    // i=1：.alx.txt，i=2：.aly.txt
                                    if (i == 0)
                                    {
                                        //X 點資料，X小於等於有效徑，才將offset的值放進陣列
                                        if (Math.Abs(double.Parse(splitPointValue[0])) <= double.Parse(effectivePath))
                                        {
                                            Array.Resize(ref arrayPointX, arrayPointX.Length + 1);
                                            arrayPointX[xCount] = splitPointValue[3];
                                            xCount++;
                                        }
                                    }
                                    else if (i == 1)
                                    {
                                        //Y 點資料，Y小於等於有效徑，才將offset的值放進陣列
                                        if (Math.Abs(double.Parse(splitPointValue[1])) <= double.Parse(effectivePath))
                                        {
                                            Array.Resize(ref arrayPointY, arrayPointY.Length + 1);
                                            arrayPointY[yCount] = splitPointValue[3];
                                            yCount++;
                                        }
                                    }
                                }
                            }
                            count++;
                        }
                    }
                    #endregion

                    #region 從陣列arrayPointX、arrayPointY 取得最大值、最小值、平均值
                    if (QcClassId == 7)
                    {
                        //白物鏡片，只取邊境 x-y (arrayPointX[0] - arrayPointY[0]、arrayPointX[xCount - 1] - arrayPointY[yCount - 1]) 的最大值
                        double firstValue = 0;
                        double finalValue = 0;

                        firstValue = Math.Abs(double.Parse(arrayPointX[0]) - double.Parse(arrayPointY[0])) * 1000;
                        finalValue = Math.Abs(double.Parse(arrayPointX[xCount - 1]) - double.Parse(arrayPointY[yCount - 1])) * 1000;

                        if (firstValue > finalValue)
                            maxmValue = firstValue;
                        else
                            maxmValue = finalValue;
                    }
                    else if (QcClassId == 8)
                    {
                        //模造鏡片，取 x-y 最大值
                        int endPoint = 0;
                        int CenterOffset = 0;
                        double sumValue = 0;

                        if (xCount >= yCount) //取最少的點資料
                            endPoint = yCount;
                        else
                            endPoint = xCount;

                        CenterOffset = Math.Abs((xCount - yCount) / 2); //計算xy中心點偏差值，要除以2取整數，無條件捨去，

                        for (int i = 0; i < endPoint; i++)
                        {
                            if (xCount >= yCount)
                                CalcValue = Math.Abs(double.Parse(arrayPointX[i + CenterOffset]) - double.Parse(arrayPointY[i]) * 1000);
                            else
                                CalcValue = Math.Abs(double.Parse(arrayPointX[i]) - double.Parse(arrayPointY[i + CenterOffset]) * 1000);

                            sumValue = sumValue + CalcValue;

                            if (CalcValue > maxmValue) //最大值
                                maxmValue = CalcValue;
                            if (CalcValue < miniValue) //最小值
                                miniValue = CalcValue;
                        }
                        averageValue = sumValue / endPoint; //平均值                        
                    }
                    #endregion
                    returnMeasureValue[0] = maxmValue.ToString();
                    returnMeasureValue[1] = miniValue.ToString();
                    returnMeasureValue[2] = averageValue.ToString();

                    return returnMeasureValue;
                }
                #endregion

                #region 解析 .csv 文件：RA Decenter
                string[] GetUa3pValueCsv(string readTextLine)
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");

                    string measureValueTilt = "";
                    string measureValueDecenter = "";
                    string measureValueRa = "";
                    string[] returnMeasureValue = new string[3];

                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        #region //.csv 檔案 包含 RA,DECENTER
                        if (readLine.IndexOf(",") != -1)
                        {
                            string[] splitReadLine = readLine.Split(',');
                            if (splitReadLine[0].IndexOf("Ra") != -1)
                            {
                                measureValueRa = splitReadLine[1];
                            }
                            else if (splitReadLine[0].IndexOf("Decenter") != -1)
                            {
                                measureValueDecenter = splitReadLine[1];
                            }
                            else if (splitReadLine[0].IndexOf("Tilt") != -1)
                            {
                                measureValueTilt = (double.Parse(splitReadLine[1]) * 60 + double.Parse(splitReadLine[3])).ToString();
                            }
                        }
                        #endregion
                        count++;
                    }
                    returnMeasureValue[0] = measureValueDecenter; //Decenter
                    returnMeasureValue[1] = measureValueTilt; //Tilt
                    returnMeasureValue[2] = measureValueRa; //Tilt
                    return returnMeasureValue;
                }
                #endregion

                #region 解析矢高 excel sag.xlsx
                #region 取得量測值
                string GetExcel(string filePath)
                {
                    //a101、a128
                    XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                    IXLWorksheet excelsheet = excelWorkbook.Worksheet(1);
                    var firstCell = excelsheet.FirstCellUsed();
                    var lastCell = excelsheet.LastCellUsed();
                    var rows = excelsheet.RangeUsed().RowsUsed();
                    var columns = excelsheet.RangeUsed().ColumnsUsed();
                    string serialCodeSagittal = "";
                    int countSagittal = 0;
                    var caveNo = "";
                    double measureValue;
                    var sagittala = "";
                    var sagittalb = "";
                    int dataColumnNumber = 0;

                    for (int i = 1; i < 3; i++) //R1、R2
                    {
                        //R1 A欄CAVE、B欄矢高、D欄矢高平均
                        //R2 F欄CAVE、G欄矢高、I欄矢高平均
                        foreach (var dataRow in rows)
                        {
                            int dataRowNumber = dataRow.RowNumber();
                            if (dataRowNumber > 5)
                            {
                                countSagittal++;
                                var cella = "";

                                if (i == 1)
                                {
                                    dataColumnNumber = 2;
                                    cella = dataRow.Cell(1).Value.ToString(); //cave merge
                                    serialCodeSagittal = "a1" + string.Format("{0:00}", 1);
                                }
                                else
                                {
                                    dataColumnNumber = 7;
                                    cella = dataRow.Cell(6).Value.ToString(); //cave merge
                                    serialCodeSagittal = "a1" + string.Format("{0:00}", 28);
                                }

                                if (cella != "") //cave is merge, 抓兩列的數據，平均
                                {
                                    caveNo = cella;
                                    sagittala = "";
                                    sagittalb = "";
                                    sagittala = dataRow.Cell(dataColumnNumber).Value.ToString();
                                }
                                else
                                {
                                    sagittalb = dataRow.Cell(dataColumnNumber).Value.ToString();
                                }

                                if (sagittala != "" && sagittalb != "") //第一、二列為不為空，才運算
                                {
                                    measureValue = (Convert.ToDouble(sagittala) + Convert.ToDouble(sagittalb)) / 2;

                                    if (measureValue != 0)
                                    {
                                        GetQcItemPrinciple(serialCodeSagittal, caveNo, measureValue.ToString(), MachineNumber[fileCount], MachineName[fileCount], UserNo[fileCount]);
                                    }
                                }
                            }
                        }
                    }
                    //資料上傳完畢後，刪除儲存的檔案
                    System.IO.File.Delete(filePath);
                    return "done";
                }

                #endregion
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModels.Add(resolveFileResultModel);
                #endregion

                logger.Error(e.Message);
            }

            return resolveFileResultModels;
        }
        #endregion

        #region //ResolvesFileOfCmmZeiss 解析蔡司三次元文件
        public List<ResolveFileResultModel> ResolvesFileOfCmmZeiss(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModels = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            #region 取得檔案及刻號 chr_#*.txt
                            string letteringNo = "";
                            string filename = item["FileFullName"].ToString();
                            if (filename.Contains("chr") == true && filename.Split('_').Length > 3)
                            {
                                string subFileId = filename.Substring(filename.IndexOf("chr"));
                                letteringNo = subFileId.Split('_')[1].Replace(".txt", string.Empty);
                            }
                            else
                            {
                                throw new SystemException("請確認上傳的檔案是否正確：" + filename);
                            }
                            if (letteringNo != "")
                            {
                                byte[] fileBytes = (byte[])item["FileContent"];
                                string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                                GetZeissTxt(readTextLine, letteringNo);
                            }
                            else
                            {
                                throw new SystemException("上傳的檔案名稱查無刻號：" + filename);
                            }
                            #endregion

                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region 解析Zeiss相關

                #region 取得量測值
                string GetZeissTxt(string readTextLine = "", string letteringNo = "")
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");
                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        string measureItem = "";
                        string measureValue = "";
                        double zAxis;
                        double design;
                        double upper;
                        double lower;
                        string[] splitReadLine = readLine.Split('\t');

                        if (splitReadLine[0] == "planid" || splitReadLine[0] == "END")
                        {
                            //第一行標題列跳過
                            count++;
                            continue;
                        }
                        else
                        {
                            measureItem = splitReadLine[2];
                            measureValue = splitReadLine[5];
                            string spString6 = splitReadLine[6]; //design
                            string spString7 = splitReadLine[7]; //upper
                            string spString8 = splitReadLine[8]; //lower
                            string spString33 = splitReadLine[33]; //zaxis

                            //轉型：string 轉 double ，並判斷，若無法轉，賦值為0
                            if (!double.TryParse(spString6, out design))
                                design = 0D;
                            if (!double.TryParse(spString7, out upper))
                                upper = 0D;
                            if (!double.TryParse(spString8, out lower))
                                lower = 0D;
                            if (!double.TryParse(spString33, out zAxis))
                                zAxis = 0D;

                            if (splitReadLine[2].IndexOf('_') != -1)
                            {
                                measureItem = splitReadLine[2].Split('_')[0] + '_' + splitReadLine[2].Split('_')[1];
                            }

                            GetQcItemPrinciple(measureItem, letteringNo, measureValue, zAxis, "", design, upper, lower);
                        }
                    }
                    return "done";
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string principleDesc = "", string letteringNo = "", string measureValue = "", double zAxis = -1, string relationshipItem = "", double design = -1, double upper = -1, double lower = -1)
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, "", principleDesc, QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["PrincipleDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MachineNumber = Machinenumber,
                            MachineName = Machinename,
                            MeasureValue = measureValue,
                            QcItemDesc = qcItemDesc,
                            ZAxis = zAxis,
                            DesignValue = design,
                            UpperTolerance = upper,
                            LowerTolerance = lower

                        };

                        resolveFileResultModels.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion

                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModels.Add(resolveFileResultModel);
                #endregion

                logger.Error(e.Message);
            }
            return resolveFileResultModels;
        }
        #endregion

        #region //ResolvesFileOfCmm3M 解析3M三次元文件
        public List<ResolveFileResultModel> ResolvesFileOfCmm3M(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModels = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            #region 取得檔案及刻號 *.txt
                            string letteringNo = "";
                            string filename = item["FileFullName"].ToString();
                            if (filename.Split('_').Length > 1)
                            {
                                letteringNo = filename.Split('_')[1].Replace(".csv", string.Empty);
                            }
                            else
                            {
                                throw new SystemException("請確認上傳的檔案是否正確：" + filename);
                            }
                            if (letteringNo != "")
                            {
                                byte[] fileBytes = (byte[])item["FileContent"];
                                string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                                Get3MTxt(readTextLine, letteringNo);
                            }
                            else
                            {
                                throw new SystemException("上傳的檔案名稱查無刻號：" + filename);
                            }
                            #endregion

                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region 解析3M相關

                #region 取得量測值
                string Get3MTxt(string readTextLine = "", string letteringNo = "")
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");
                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        string measureItem = "";
                        string measureValue = "";
                        string[] splitReadLine = readLine.Split(';');

                        if (splitReadLine[2] == "")
                            measureItem = splitReadLine[1];
                        else
                            measureItem = splitReadLine[1] + "_" + splitReadLine[2];

                        measureValue = splitReadLine[6];
                        GetQcItemPrinciple(measureItem, letteringNo, measureValue);
                    }
                    return "done";
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string principleDesc = "", string letteringNo = "", string measureValue = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, "", principleDesc, QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["PrincipleDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MachineNumber = Machinenumber,
                            MachineName = Machinename,
                            MeasureValue = measureValue,
                            QcItemDesc = qcItemDesc,

                        };

                        resolveFileResultModels.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion

                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModels.Add(resolveFileResultModel);
                #endregion

                logger.Error(e.Message);
            }
            return resolveFileResultModels;
        }
        #endregion

        #region //ResolvesFileOfZYGO 解析ZYGO文件
        public List<ResolveFileResultModel> ResolvesFileOfZYGO(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModals = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            string filePath = @"\\192.168.20.199\d8-系統開發室\99-單位共用區\滕峮\檔案測試\ZYGOTemp.xlsx";
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            byte[] fileBytes = (byte[])item["FileContent"];
                            string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes); //BIG5
                            //將byte[] 儲存至預設路徑 filePath
                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            //資料解析
                            GeZYGOExcel();
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region 解析ZYGO相關

                #region 取得量測值
                string GeZYGOExcel()
                {
                    XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                    IXLWorksheet excelsheet = excelWorkbook.Worksheet(1);
                    var firstCell = excelsheet.FirstCellUsed();
                    var lastCell = excelsheet.LastCellUsed();
                    var rows = excelsheet.RangeUsed().RowsUsed();
                    var columns = excelsheet.RangeUsed().ColumnsUsed();

                    foreach (var dataRow in rows)
                    {
                        string serialCode = "";
                        string zygoNo = "";
                        string letteringNo = "";
                        string measureValue = "";
                        int siteCount = 0;

                        if (dataRow.RowNumber() > 3)
                        {
                            var row = excelsheet.Row(dataRow.RowNumber());
                            zygoNo = row.Cell(1).Value.ToString();

                            if (Regex.IsMatch(zygoNo, @"^[0-9]") && zygoNo.IndexOf('/') == -1 && zygoNo.IndexOf("Sigma") == -1 && zygoNo != string.Empty && zygoNo.IndexOf('-') == -1 && zygoNo.IndexOf('~') == -1)
                            {
                                if (Regex.IsMatch(zygoNo, @"[0-9]G"))
                                {
                                    serialCode = "p001"; //"噴砂粗糙度-澆口"
                                    letteringNo = zygoNo.Replace("G", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]Y"))
                                {
                                    serialCode = "p003"; //"噴砂粗糙度-外"
                                    letteringNo = zygoNo.Replace("Y", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]L"))
                                {
                                    serialCode = "p004"; //"噴砂粗糙度-直身段"
                                    letteringNo = zygoNo.Replace("L", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]A"))
                                {
                                    serialCode = "p005"; //"噴砂粗糙度-內上 0度"
                                    letteringNo = zygoNo.Replace("A", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YA"))
                                {
                                    serialCode = "p009"; //"噴砂粗糙度-外上 0度"
                                    letteringNo = zygoNo.Replace("YA", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]B"))
                                {
                                    serialCode = "p006"; //"噴砂粗糙度-內下 180度"
                                    letteringNo = zygoNo.Replace("B", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YB"))
                                {
                                    serialCode = "p010"; //"噴砂粗糙度-外下 180度"
                                    letteringNo = zygoNo.Replace("YB", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]C"))
                                {
                                    serialCode = "p007"; //"噴砂粗糙度-內左 90度"
                                    letteringNo = zygoNo.Replace("C", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YC"))
                                {
                                    serialCode = "p011"; //"噴砂粗糙度-外左 90度"
                                    letteringNo = zygoNo.Replace("YC", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]D"))
                                {
                                    serialCode = "p008"; //"噴砂粗糙度-內右 270度"
                                    letteringNo = zygoNo.Replace("D", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YD"))
                                {
                                    serialCode = "p012"; //"噴砂粗糙度-外右 270度"
                                    letteringNo = zygoNo.Replace("YD", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"S[0-9]")) //斜面
                                {
                                    siteCount = Convert.ToInt16(zygoNo.Split('S')[1]);
                                    letteringNo = zygoNo.Split('S')[0];
                                    serialCode = "p0" + (siteCount + 12).ToString("00");
                                }
                                else if (Regex.IsMatch(zygoNo, @"PV[0-9]")) //pv
                                {
                                    siteCount = Convert.ToInt16(zygoNo.Split('V')[1]);
                                    letteringNo = zygoNo.Split('P')[0];
                                    serialCode = "l0" + siteCount.ToString("00");
                                }
                                else if (Regex.IsMatch(zygoNo, @"SV[0-9]")) //斜面pv
                                {
                                    siteCount = Convert.ToInt16(zygoNo.Split('V')[1]);
                                    letteringNo = zygoNo.Split('S')[0];
                                    serialCode = "l0" + (siteCount + 4).ToString("00");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]+"))
                                {
                                    serialCode = "p002"; //"噴砂粗糙度-內"
                                    letteringNo = zygoNo;
                                }

                                if (serialCode != string.Empty)
                                {
                                    measureValue = row.Cell(3).Value.ToString();
                                    GetQcItemPrinciple(serialCode, letteringNo, measureValue);
                                }
                            }
                        }

                    }
                    return "done";
                }

                string GeZYGOTxt(string readTextLine = "")
                {
                    string[] readTextLineArray = Regex.Split(readTextLine, "\r\n");
                    short count = 0;
                    foreach (var readLine in readTextLineArray)
                    {
                        #region //忽略空白
                        if (readLine == "")
                        {
                            count++;
                            continue;
                        }
                        #endregion

                        string[] splitReadLine = readLine.Split(',');
                        string letteringNo = "";
                        string zygoNo = splitReadLine[0];
                        string serialCode = "";

                        int siteCount = 0;
                        string measureValue = "";

                        if (zygoNo == "Statistic")
                        {
                            break;
                        }
                        if (Regex.IsMatch(zygoNo, @"^[0-9]") && zygoNo != string.Empty && zygoNo.IndexOf('-') == -1 && zygoNo.IndexOf('~') == -1)
                        {
                            if (Regex.IsMatch(zygoNo, @"[0-9]G"))
                            {
                                serialCode = "p001"; //"噴砂粗糙度-澆口"
                                letteringNo = zygoNo.Replace("G", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]Y"))
                            {
                                serialCode = "p003"; //"噴砂粗糙度-外"
                                letteringNo = zygoNo.Replace("Y", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]L"))
                            {
                                serialCode = "p004"; //"噴砂粗糙度-直身段"
                                letteringNo = zygoNo.Replace("L", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]A"))
                            {
                                serialCode = "p005"; //"噴砂粗糙度-內上 0度"
                                letteringNo = zygoNo.Replace("A", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]YA"))
                            {
                                serialCode = "p009"; //"噴砂粗糙度-外上 0度"
                                letteringNo = zygoNo.Replace("YA", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]B"))
                            {
                                serialCode = "p006"; //"噴砂粗糙度-內下 180度"
                                letteringNo = zygoNo.Replace("B", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]YB"))
                            {
                                serialCode = "p010"; //"噴砂粗糙度-外下 180度"
                                letteringNo = zygoNo.Replace("YB", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]C"))
                            {
                                serialCode = "p007"; //"噴砂粗糙度-內左 90度"
                                letteringNo = zygoNo.Replace("C", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]YC"))
                            {
                                serialCode = "p011"; //"噴砂粗糙度-外左 90度"
                                letteringNo = zygoNo.Replace("YC", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]D"))
                            {
                                serialCode = "p008"; //"噴砂粗糙度-內右 270度"
                                letteringNo = zygoNo.Replace("D", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]YD"))
                            {
                                serialCode = "p012"; //"噴砂粗糙度-外右 270度"
                                letteringNo = zygoNo.Replace("YD", "");
                            }
                            else if (Regex.IsMatch(zygoNo, @"S[0-9]")) //斜面
                            {
                                siteCount = Convert.ToInt16(zygoNo.Split('S')[1]);
                                letteringNo = zygoNo.Split('S')[0];
                                serialCode = "p0" + (siteCount + 12).ToString("00");
                            }
                            else if (Regex.IsMatch(zygoNo, @"PV[0-9]")) //pv
                            {
                                siteCount = Convert.ToInt16(zygoNo.Split('V')[1]);
                                letteringNo = zygoNo.Split('P')[0];
                                serialCode = "l0" + siteCount.ToString("00");
                            }
                            else if (Regex.IsMatch(zygoNo, @"SV[0-9]")) //斜面pv
                            {
                                siteCount = Convert.ToInt16(zygoNo.Split('V')[1]);
                                letteringNo = zygoNo.Split('S')[0];
                                serialCode = "l0" + (siteCount + 4).ToString("00");
                            }
                            else if (Regex.IsMatch(zygoNo, @"[0-9]+"))
                            {
                                serialCode = "p002"; //"噴砂粗糙度-內"
                                letteringNo = zygoNo;
                            }


                            if (serialCode != string.Empty)
                            {
                                measureValue = splitReadLine[1];
                                GetQcItemPrinciple(serialCode, letteringNo, measureValue);
                            }
                        }
                    }
                    return "done";
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string serialCode = "", string letteringNo = "", string measureValue = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["PrincipleDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MeasureValue = measureValue,
                            QcItemDesc = qcItemDesc,
                            MachineNumber = Machinenumber,
                            MachineName = Machinename
                        };

                        resolveFileResultModals.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion

                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModal = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModals.Add(resolveFileResultModal);
                #endregion

                logger.Error(e.Message);
            }
            //資料上傳完畢後，刪除儲存的檔案
            System.IO.File.Delete(filePath);
            return resolveFileResultModals;
        }
        #endregion

        #region //ResolvesFileOfZYGOforJC 解析ZYGO文件 for 晶彩
        public List<ResolveFileResultModel> ResolvesFileOfZYGOforJC(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModals = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            string filePath = @"\\192.168.20.199\d8-系統開發室\99-單位共用區\滕峮\檔案測試\ZYGOTemp.xlsx";
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            byte[] fileBytes = (byte[])item["FileContent"];
                            string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes); //BIG5
                            //將byte[] 儲存至預設路徑 filePath
                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            //資料解析
                            GeZYGOExcel();
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region 解析ZYGO相關

                #region 取得量測值
                string GeZYGOExcel()
                {
                    XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                    IXLWorksheet excelsheet = excelWorkbook.Worksheet(1);
                    var firstCell = excelsheet.FirstCellUsed();
                    var lastCell = excelsheet.LastCellUsed();
                    var rows = excelsheet.RangeUsed().RowsUsed();
                    var columns = excelsheet.RangeUsed().ColumnsUsed();

                    foreach (var dataRow in rows)
                    {
                        string serialCode = "";
                        string zygoNo = "";
                        string letteringNo = "";
                        string measureValue = "";
                        int siteCount = 0;

                        if (dataRow.RowNumber() > 3)
                        {
                            var row = excelsheet.Row(dataRow.RowNumber());
                            zygoNo = row.Cell(1).Value.ToString();

                            if (Regex.IsMatch(zygoNo, @"^[0-9]") && zygoNo.IndexOf('/') == -1 && zygoNo.IndexOf("Sigma") == -1 && zygoNo != string.Empty && zygoNo.IndexOf('-') == -1 && zygoNo.IndexOf('~') == -1)
                            {
                                if (Regex.IsMatch(zygoNo, @"[0-9]G"))
                                {
                                    serialCode = "p001"; //"噴砂粗糙度-澆口"
                                    letteringNo = zygoNo.Replace("G", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]Y"))
                                {
                                    serialCode = "p003"; //"噴砂粗糙度-外"
                                    letteringNo = zygoNo.Replace("Y", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]L"))
                                {
                                    serialCode = "p004"; //"噴砂粗糙度-直身段"
                                    letteringNo = zygoNo.Replace("L", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]A"))
                                {
                                    serialCode = "p005"; //"噴砂粗糙度-內上 0度"
                                    letteringNo = zygoNo.Replace("A", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YA"))
                                {
                                    serialCode = "p009"; //"噴砂粗糙度-外上 0度"
                                    letteringNo = zygoNo.Replace("YA", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]B"))
                                {
                                    serialCode = "p006"; //"噴砂粗糙度-內下 180度"
                                    letteringNo = zygoNo.Replace("B", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YB"))
                                {
                                    serialCode = "p010"; //"噴砂粗糙度-外下 180度"
                                    letteringNo = zygoNo.Replace("YB", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]C"))
                                {
                                    serialCode = "p007"; //"噴砂粗糙度-內左 90度"
                                    letteringNo = zygoNo.Replace("C", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YC"))
                                {
                                    serialCode = "p011"; //"噴砂粗糙度-外左 90度"
                                    letteringNo = zygoNo.Replace("YC", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]D"))
                                {
                                    serialCode = "p008"; //"噴砂粗糙度-內右 270度"
                                    letteringNo = zygoNo.Replace("D", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]YD"))
                                {
                                    serialCode = "p012"; //"噴砂粗糙度-外右 270度"
                                    letteringNo = zygoNo.Replace("YD", "");
                                }
                                else if (Regex.IsMatch(zygoNo, @"S[0-9]")) //斜面
                                {
                                    siteCount = Convert.ToInt16(zygoNo.Split('S')[1]);
                                    letteringNo = zygoNo.Split('S')[0];
                                    serialCode = "p0" + (siteCount + 12).ToString("00");
                                }
                                else if (Regex.IsMatch(zygoNo, @"PV[0-9]")) //pv
                                {
                                    siteCount = Convert.ToInt16(zygoNo.Split('V')[1]);
                                    letteringNo = zygoNo.Split('P')[0];
                                    serialCode = "l0" + siteCount.ToString("00");
                                }
                                else if (Regex.IsMatch(zygoNo, @"SV[0-9]")) //斜面pv
                                {
                                    siteCount = Convert.ToInt16(zygoNo.Split('V')[1]);
                                    letteringNo = zygoNo.Split('S')[0];
                                    serialCode = "l0" + (siteCount + 4).ToString("00");
                                }
                                else if (Regex.IsMatch(zygoNo, @"[0-9]+"))
                                {
                                    serialCode = "p002"; //"噴砂粗糙度-內"
                                    letteringNo = zygoNo;
                                }

                                if (serialCode != string.Empty)
                                {
                                    measureValue = row.Cell(2).Value.ToString();
                                    GetQcItemPrinciple(serialCode, letteringNo, measureValue);
                                }
                            }
                        }

                    }
                    return "done";
                }

                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string serialCode = "", string letteringNo = "", string measureValue = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["PrincipleDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo != "")
                    {
                        if (letterNoList.IndexOf(letteringNo) == -1)
                        {
                            letterNoList.Add(letteringNo);
                        }

                        ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                        {
                            QcItemNo = principleFullNo,
                            QcItemName = qcItemName,
                            LetteringNo = letteringNo,
                            MeasureValue = measureValue,
                            QcItemDesc = qcItemDesc,
                            MachineNumber = Machinenumber,
                            MachineName = Machinename
                        };

                        resolveFileResultModals.Add(resolveFileResultModel);
                    }

                    return principleFullNo;
                }
                #endregion

                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModal = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModals.Add(resolveFileResultModal);
                #endregion

                logger.Error(e.Message);
            }
            //資料上傳完畢後，刪除儲存的檔案
            System.IO.File.Delete(filePath);
            return resolveFileResultModals;
        }
        #endregion

        #region //ResolvesFileOfGeneric 解析公版excel文件 .xlsx
        public List<ResolveFileResultModel> ResolvesFileOfGeneric(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModals = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            string filePath = @"\\192.168.20.199\d8-系統開發室\99-單位共用區\滕峮\檔案測試\GenericTemp.xlsx";
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            byte[] fileBytes = (byte[])item["FileContent"];
                            string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes); //BIG5
                            ////將byte[] 儲存至預設路徑 filePath
                            System.IO.File.WriteAllBytes(filePath, fileBytes);
                            //資料解析
                            GetExcel();
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region 解析excel相關

                #region 取得量測值
                string GetExcel()
                {
                    XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                    IXLWorksheet excelsheet = excelWorkbook.Worksheet(1);
                    var firstCell = excelsheet.FirstCellUsed();
                    var lastCell = excelsheet.LastCellUsed();
                    var rows = excelsheet.RangeUsed().RowsUsed();
                    var columns = excelsheet.RangeUsed().ColumnsUsed();

                    foreach (var dataRow in rows)
                    {
                        string serialCode = "";
                        string letteringNo = "";
                        string measureValue = "";
                        string ItemName = "";
                        string ItemRemark = "";
                        double design;
                        double upper;
                        double lower;
                        double zAxis;
                        string userno;

                        var row = excelsheet.Row(dataRow.RowNumber());

                        if (dataRow.RowNumber() != 1)
                        {
                            string cellMachinename = row.Cell(4).Value.ToString();

                            var cell1 = row.Cell(1).Value.ToString();
                            if (cellMachinename != "")
                            {
                                serialCode = cell1.Substring(0, 3) + cell1.Substring(6);
                                Machinenumber = cell1.Substring(3, 3);
                            }
                            else
                            {
                                serialCode = row.Cell(1).Value.ToString();
                            }

                            ItemName = row.Cell(2).Value.ToString();
                            ItemRemark = row.Cell(3).Value.ToString();
                            Machinename = row.Cell(4).Value.ToString();
                            string cell5 = row.Cell(5).Value.ToString(); //design
                            string cell6 = row.Cell(6).Value.ToString(); //upper
                            string cell7 = row.Cell(7).Value.ToString(); //lower
                            string cell8 = row.Cell(8).Value.ToString(); //zaxis
                            userno = row.Cell(9).Value.ToString();

                            foreach (var dataColumn in columns)
                            {
                                int dataColumnNumber = dataColumn.ColumnNumber();
                                if (dataColumnNumber > 10)
                                {
                                    var column = excelsheet.Column(dataColumnNumber);
                                    letteringNo = excelsheet.Row(1).Cell(dataColumnNumber).Value.ToString();

                                    measureValue = row.Cell(dataColumnNumber).Value.ToString();

                                    //轉型：string 轉 double ，並判斷，若無法轉，賦值為0
                                    if (!double.TryParse(cell5, out design))
                                        design = 0D;
                                    if (!double.TryParse(cell6, out upper))
                                        upper = 0D;
                                    if (!double.TryParse(cell7, out lower))
                                        lower = 0D;
                                    if (!double.TryParse(cell8, out zAxis))
                                        zAxis = 0D;

                                    GetQcItemPrinciple(userno, serialCode, ItemName, ItemRemark, letteringNo, measureValue, zAxis, design, upper, lower);

                                }
                            }
                        }

                    }
                    return "done";
                }

                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string Userno = "", string serialCode = "", string qcItemName = "", string qcItemDesc = "", string letteringNo = "", string measureValue = "", double zAxis = -1, double design = -1, double upper = -1, double lower = -1)
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    if (Userno == "")
                        Userno = null;

                    string principleFullNo = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["QcItemDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo == "")
                    {
                        principleFullNo = serialCode;
                    }
                    if (letterNoList.IndexOf(letteringNo) == -1)
                    {
                        letterNoList.Add(letteringNo);
                    }

                    ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                    {
                        QcItemNo = principleFullNo,
                        QcItemName = qcItemName,
                        LetteringNo = letteringNo,
                        MachineNumber = Machinenumber,
                        MachineName = Machinename,
                        MeasureValue = measureValue,
                        QcItemDesc = qcItemDesc,
                        DesignValue = design,
                        UpperTolerance = upper,
                        LowerTolerance = lower,
                        UserNo = Userno,
                        ZAxis = Convert.ToDouble(zAxis)
                    };

                    resolveFileResultModals.Add(resolveFileResultModel);

                    return principleFullNo;
                }
                #endregion

                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModal = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModals.Add(resolveFileResultModal);
                #endregion

                logger.Error(e.Message);
            }

            //資料上傳完畢後，刪除儲存的檔案
            System.IO.File.Delete(filePath);
            return resolveFileResultModals;
        }
        #endregion

        #region //ResolvesFileOfTR 解析TR偏芯文件 .xlsx for 中揚
        public List<ResolveFileResultModel> ResolvesFileOfTR(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModals = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            string filePath = @"\\192.168.20.199\d8-系統開發室\99-單位共用區\滕峮\檔案測試\TRTemp.xlsx";
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            byte[] fileBytes = (byte[])item["FileContent"];
                            string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                            ////將byte[] 儲存至預設路徑 filePath
                            System.IO.File.WriteAllBytes(filePath, fileBytes);

                            #region //解析excel相關
                            //k001、k101、k301、k601、k701
                            XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                            IXLWorksheet excelsheet = excelWorkbook.Worksheet(1);
                            var firstCell = excelsheet.FirstCellUsed();
                            var lastCell = excelsheet.LastCellUsed();
                            var rows = excelsheet.RangeUsed().RowsUsed();
                            var columns = excelsheet.RangeUsed().ColumnsUsed();
                            var ietteringNoRow = excelsheet.Row(28);
                            int col = 2;

                            for (int i = 1; i <= 24; i++)
                            {
                                if (i == 13) col = 10; //換欄
                                var eachRows = excelsheet.Row(i + 2);
                                string cellSite = eachRows.Cell(col + 1).Value.ToString();

                                //編號有合併儲存格，要跳過第二行

                                int letteringNo = 0;
                                if (i % 2 == 0) letteringNo = i / 2; else letteringNo = i / 2 + 1;

                                string tiltSerial = "";
                                string decenterSerial = "k101";
                                string absSerial = "";
                                string XaxisSerial = "";
                                string YaxisSerial = "";

                                if (cellSite == "R1")
                                {
                                    tiltSerial = "k001";
                                    absSerial = "k301";
                                    XaxisSerial = "k601";
                                    YaxisSerial = "k701";
                                }
                                else if (cellSite == "R2")
                                {
                                    tiltSerial = "k002";
                                    absSerial = "k302";
                                    XaxisSerial = "k602";
                                    YaxisSerial = "k702";
                                }

                                //偏芯absolute(反射偏芯)、X、Y
                                var XaxisValue = eachRows.Cell(col + 2).CachedValue.ToString();
                                if (XaxisValue != "") GetQcItemPrinciple(XaxisSerial, letteringNo.ToString(), XaxisValue);

                                var YaxisValue = eachRows.Cell(col + 3).CachedValue.ToString();
                                if (YaxisValue != "") GetQcItemPrinciple(YaxisSerial, letteringNo.ToString(), YaxisValue);

                                var absValue = eachRows.Cell(col + 4).CachedValue.ToString();
                                if (absValue != "") GetQcItemPrinciple(absSerial, letteringNo.ToString(), absValue);

                                //Decenter 有合併儲存格，要跳過第二行
                                if (i % 2 != 0)
                                {
                                    var decenterValue = eachRows.Cell(col + 5).CachedValue.ToString();
                                    if (decenterValue != "") GetQcItemPrinciple(decenterSerial, letteringNo.ToString(), decenterValue);
                                }

                                //tilt
                                var tiltValue = eachRows.Cell(col + 6).CachedValue.ToString();
                                if (tiltValue != "") GetQcItemPrinciple(tiltSerial, letteringNo.ToString(), tiltValue);
                            }
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string serialCode = "", string letteringNo = "", string measureValue = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["QcItemDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo == "")
                    {
                        principleFullNo = serialCode;
                    }
                    if (letterNoList.IndexOf(letteringNo) == -1)
                    {
                        letterNoList.Add(letteringNo);
                    }

                    ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                    {
                        QcItemNo = principleFullNo,
                        QcItemName = qcItemName,
                        LetteringNo = letteringNo,
                        MachineNumber = Machinenumber,
                        MachineName = Machinename,
                        MeasureValue = measureValue,
                        QcItemDesc = qcItemDesc,
                    };

                    resolveFileResultModals.Add(resolveFileResultModel);

                    return principleFullNo;
                }
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModal = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModals.Add(resolveFileResultModal);
                #endregion

                logger.Error(e.Message);
            }

            //資料上傳完畢後，刪除儲存的檔案
            System.IO.File.Delete(filePath);
            return resolveFileResultModals;
        }
        #endregion

        #region //ResolvesFileOfTRforJC 解析TR偏芯文件 .xlsx for 晶彩
        public List<ResolveFileResultModel> ResolvesFileOfTRforJC(int FileId = -1, int QcClassId = -1, int QmmDetailId = -1, string QcEncoding = "", string Machinenumber = "", string Machinename = "")
        {
            List<ResolveFileResultModel> resolveFileResultModals = new List<ResolveFileResultModel>();
            List<string> errorQcItemList = new List<string>();
            List<string> letterNoList = new List<string>();
            JObject dataRequestJson = new JObject();
            string filePath = @"\\192.168.20.199\d8-系統開發室\99-單位共用區\滕峮\檔案測試\TRTempforJC.xlsx";
            try
            {
                #region //確認檔案資料是否正確
                List<string> FileIdList = new List<string>();
                FileIdList.Add(FileId.ToString());
                string FileIdListString = String.Join(", ", FileIdList.ToArray());
                dataRequest = qcDataManagementDA.GetFileInfo(FileIdListString, "");
                dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequestJson["status"].ToString() == "success")
                {
                    if (dataRequestJson["data"].ToString() != "[]")
                    {
                        foreach (var item in dataRequestJson["data"])
                        {
                            byte[] fileBytes = (byte[])item["FileContent"];
                            string readTextLine = Encoding.GetEncoding(QcEncoding).GetString(fileBytes);
                            ////將byte[] 儲存至預設路徑 filePath
                            System.IO.File.WriteAllBytes(filePath, fileBytes);

                            #region 解析excel相關
                            //k001、k101、k301、k601、k701
                            XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                            IXLWorksheet excelsheet = excelWorkbook.Worksheet("整理");
                            var firstCell = excelsheet.FirstCellUsed();
                            var lastCell = excelsheet.LastCellUsed();
                            var rows = excelsheet.RangeUsed().RowsUsed();
                            var columns = excelsheet.RangeUsed().ColumnsUsed();
                            int col = 1;
                            string CaveNo = "";

                            foreach (var dataRows in rows)
                            {
                                int dataRowsNumber = dataRows.RowNumber();
                                if (dataRowsNumber > 5)
                                {
                                    if (dataRowsNumber > 53)
                                    {
                                        break;
                                    }

                                    if ((dataRowsNumber - 5) % 2 != 0) CaveNo = dataRows.Cell(col).Value.ToString().Trim();

                                    string cellSite = dataRows.Cell(col + 1).Value.ToString().Trim();

                                    string tiltSerial = "";
                                    string decenterSerial = "k101";
                                    string absSerial = "";
                                    string XaxisSerial = "";
                                    string YaxisSerial = "";

                                    if (cellSite == "R1")
                                    {
                                        tiltSerial = "k001";
                                        absSerial = "k301";
                                        XaxisSerial = "k601";
                                        YaxisSerial = "k701";
                                    }
                                    else if (cellSite == "R2")
                                    {
                                        tiltSerial = "k002";
                                        absSerial = "k302";
                                        XaxisSerial = "k602";
                                        YaxisSerial = "k702";
                                    }

                                    //偏芯absolute(反射偏芯)、X、Y
                                    var XaxisValue = dataRows.Cell(col + 2).CachedValue.ToString().Trim();
                                    if (XaxisValue != "") GetQcItemPrinciple(XaxisSerial, CaveNo, XaxisValue);

                                    var YaxisValue = dataRows.Cell(col + 3).CachedValue.ToString().Trim();
                                    if (YaxisValue != "") GetQcItemPrinciple(YaxisSerial, CaveNo, YaxisValue);

                                    var absValue = dataRows.Cell(col + 5).CachedValue.ToString().Trim();
                                    if (absValue != "") GetQcItemPrinciple(absSerial, CaveNo, absValue);

                                    //Decenter 有合併儲存格，要跳過第二行
                                    var decenterValue = dataRows.Cell(col + 4).CachedValue.ToString().Trim();
                                    if (decenterValue != "") GetQcItemPrinciple(decenterSerial, CaveNo, decenterValue);
                                }
                            }
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //取得量測項目編碼原則
                string GetQcItemPrinciple(string serialCode = "", string letteringNo = "", string measureValue = "")
                {
                    dataRequest = qcDataManagementDA.GetQcItemPrinciple(-1, QcClassId, serialCode, "", QmmDetailId);
                    dataRequestJson = JObject.Parse(dataRequest);

                    string principleFullNo = "";
                    string qcItemName = "";
                    string qcItemDesc = "";
                    string principleDesc = "";
                    if (dataRequestJson["status"].ToString() == "success")
                    {
                        if (dataRequestJson["data"].ToString() != "[]")
                        {
                            foreach (var item in dataRequestJson["data"])
                            {
                                if (Convert.ToInt32(item["QcItemId"]) > 0)
                                {
                                    principleFullNo = item["PrincipleFullNo"].ToString();
                                    principleDesc = item["PrincipleDesc"].ToString();
                                    qcItemName = item["QcItemName"].ToString();
                                    qcItemDesc = item["QcItemDesc"].ToString();
                                }
                                else
                                {
                                    errorQcItemList.Add(principleDesc);
                                }
                            }
                        }
                        else
                        {
                            errorQcItemList.Add(principleDesc);
                        }
                    }

                    if (principleFullNo == "")
                    {
                        principleFullNo = serialCode;
                    }
                    if (letterNoList.IndexOf(letteringNo) == -1)
                    {
                        letterNoList.Add(letteringNo);
                    }

                    ResolveFileResultModel resolveFileResultModel = new ResolveFileResultModel()
                    {
                        QcItemNo = principleFullNo,
                        QcItemName = qcItemName,
                        LetteringNo = letteringNo,
                        MachineNumber = Machinenumber,
                        MachineName = Machinename,
                        MeasureValue = measureValue,
                        QcItemDesc = qcItemDesc,
                    };

                    resolveFileResultModals.Add(resolveFileResultModel);

                    return principleFullNo;
                }
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                ResolveFileResultModel resolveFileResultModal = new ResolveFileResultModel()
                {
                    ErrorMessage = e.Message
                };

                resolveFileResultModals.Add(resolveFileResultModal);
                #endregion

                logger.Error(e.Message);
            }

            //資料上傳完畢後，刪除儲存的檔案
            System.IO.File.Delete(filePath);
            return resolveFileResultModals;
        }
        #endregion

        #region //ReloadQcItemDateForExcel 解析Excel更新QcItem
        [HttpPost]
        public void ReloadQcItemDateForExcel(string FilePath = "", string CompanyNo = "")
        {
            try
            {
                #region //解析EXCEL
                //FilePath = "\\\\192.168.20.199\\d8-系統開發室\\99-單位共用區\\安宏\\EXCEL_TEST0315.xlsx";
                XLWorkbook workbook = new XLWorkbook(FilePath);
                IXLWorksheet worksheet = workbook.Worksheet(1);
                var firstCell = worksheet.FirstCellUsed();
                var lastCell = worksheet.LastCellUsed();

                int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                var data = worksheet.Range(firstCell.Address, lastCell.Address);
                var table = data.AsTable();

                List<QcItem> qcItems = new List<QcItem>();
                for (var i = 1; i <= CellLength; i++)
                {
                    if (table.Cell(i, 1).Value.ToString() != "")
                    {
                        QcItem qcItem = new QcItem()
                        {
                            QcItemId = Convert.ToInt32(table.Cell(i, 1).Value),
                            QcClassId = Convert.ToInt32(table.Cell(i, 2).Value),
                            QcItemNo = table.Cell(i, 3).Value.ToString(),
                            QcItemName = table.Cell(i, 4).Value.ToString(),
                            QcItemDesc = table.Cell(i, 5).Value.ToString(),
                            QcItemType = table.Cell(i, 6).Value.ToString(),
                            QcType = table.Cell(i, 7).Value.ToString(),
                            Status = table.Cell(i, 8).Value.ToString(),
                            Remark = table.Cell(i, 9).Value.ToString(),
                        };
                        qcItems.Add(qcItem);
                    }
                }
                #endregion

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.ReloadQcItemDateForExcel(qcItems, CompanyNo);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
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

        #region //ReloadQcItemPrincipleDateForExcel 解析Excel更新QcItemPrinciple
        [HttpPost]
        public void ReloadQcItemPrincipleDateForExcel(string FilePath = "", string CompanyNo = "")
        {
            int count = 0;
            try
            {
                #region //解析EXCEL
                //FilePath = "\\\\192.168.20.199\\d8-系統開發室\\99-單位共用區\\安宏\\EXCEL_TEST0315.xlsx";
                XLWorkbook workbook = new XLWorkbook(FilePath);
                IXLWorksheet worksheet = workbook.Worksheet(1);
                var firstCell = worksheet.FirstCellUsed();
                var lastCell = worksheet.LastCellUsed();

                int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                var data = worksheet.Range(firstCell.Address, lastCell.Address);
                var table = data.AsTable();

                List<QcItemPrinciple> qcItemPrinciples = new List<QcItemPrinciple>();
                for (var i = 1; i <= CellLength; i++)
                {
                    QcItemPrinciple qcItemPrinciple = new QcItemPrinciple()
                    {
                        PrincipleId = Convert.ToInt32(table.Cell(i, 1).Value),
                        QcClassId = Convert.ToInt32(table.Cell(i, 2).Value),
                        QmmDetailId = Convert.ToInt32(table.Cell(i, 3).Value),
                        PrincipleNo = table.Cell(i, 4).Value.ToString(),
                        PrincipleDesc = table.Cell(i, 5).Value.ToString()
                    };
                    qcItemPrinciples.Add(qcItemPrinciple);
                    count++;
                }
                #endregion

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.ReloadQcItemPrincipleDateForExcel(qcItemPrinciples, CompanyNo);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
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

        #region //ReloadPrincipleDetailDateForExcel 解析Excel更新PrincipleDetail
        [HttpPost]
        public void ReloadPrincipleDetailDateForExcel(string FilePath = "", string CompanyNo = "")
        {
            try
            {
                #region //解析EXCEL
                //FilePath = "\\\\192.168.20.199\\d8-系統開發室\\99-單位共用區\\安宏\\EXCEL_TEST0315.xlsx";
                XLWorkbook workbook = new XLWorkbook(FilePath);
                IXLWorksheet worksheet = workbook.Worksheet(1);
                var firstCell = worksheet.FirstCellUsed();
                var lastCell = worksheet.LastCellUsed();

                int CellLength = Convert.ToInt32(lastCell.ToString().Substring(1, lastCell.ToString().Length - 1));

                var data = worksheet.Range(firstCell.Address, lastCell.Address);
                var table = data.AsTable();

                List<PrincipleDetail> principleDetails = new List<PrincipleDetail>();
                for (var i = 1; i <= CellLength; i++)
                {
                    PrincipleDetail principleDetail = new PrincipleDetail()
                    {
                        PdId = Convert.ToInt32(table.Cell(i, 1).Value),
                        PrincipleId = Convert.ToInt32(table.Cell(i, 2).Value),
                        PrincipleDesc = table.Cell(i, 3).Value.ToString()
                    };
                    principleDetails.Add(principleDetail);
                }
                #endregion

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.ReloadPrincipleDetailDateForExcel(principleDetails, CompanyNo);
                JObject dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequestJson["status"].ToString() != "success")
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
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

        #region//Tray轉條碼
        [HttpPost]
        [Route("api/QMS/GetTrayNoToBarcodeNo")]
        public void GetTrayNoToBarcodeNo(string Company = "", string SecretKey = "", string InputNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetTrayNoToBarcodeNo");
                #endregion

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetTrayNoToBarcodeNo(Company, InputNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetAutomationDataForMTF -- 取得自動化MTF資料表數據 -- WuTC 2024-05-06
        [HttpPost]
        [Route("api/QMS/GetAutomationDataForMTF")]
        public void GetAutomationDataForMTF(string Company = "", string SecretKey = "", string AutomationRecordId = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetAutomationDataForMTF");
                #endregion

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetAutomationDataForMTF(Company, AutomationRecordId);
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
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //ResolvesQcMeasurePointData 解析量測點資料 -- WuTc 2024-11-29[HttpPost]        
        [HttpPost]
        public void ResolvesQcMeasurePointData(int QcRecordFileId = -1, string fileType = "")
        {
            try
            {
                List<QcMeasurePointData> PointDataSetList = new List<QcMeasurePointData>();
                #region //確認檔案資料是否正確
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcRecordFileID(QcRecordFileId);
                jsonResponse = JObject.Parse(dataRequest);

                if (jsonResponse["status"].ToString() == "success")
                {
                    if (jsonResponse["data"].ToString() != "[]")
                    {
                        foreach (var item in jsonResponse["data"])
                        {
                            string filePath = item["PhysicalPath"].ToString();
                            var LotNumber = item["LotNumber"].ToString();
                            var BarcodeId = item["BarcodeId"].ToString();
                            int CompanyId = Convert.ToInt16(item["CompanyId"].ToString());

                            if (filePath != "")
                            {
                                if (filePath.IndexOf("xlsx") != -1)
                                {
                                    #region //資料解析 xlsx
                                    XLWorkbook excelWorkbook = new XLWorkbook(filePath);
                                    IXLWorksheet excelsheet = excelWorkbook.Worksheet(1);
                                    var firstCell = excelsheet.FirstCellUsed();
                                    var lastCell = excelsheet.LastCellUsed();
                                    var rows = excelsheet.RangeUsed().RowsUsed();
                                    var columns = excelsheet.RangeUsed().ColumnsUsed();

                                    foreach (var dataRow in rows)
                                    {
                                        string Axis = "";
                                        string pointSite = "";
                                        string pointValue = "";
                                        string Cavity = "";
                                        int siteCount = 1;

                                        var row = excelsheet.Row(dataRow.RowNumber());
                                        int dataRowNumber = row.RowNumber();
                                        string point = row.Cell(1).Value.ToString();

                                        if (CompanyId == 4) //晶彩
                                        {
                                            if (dataRowNumber > 2) //從第幾列開始抓數據
                                            {
                                                foreach (var dataColumn in columns) //正反
                                                {
                                                    int dataColumnNumber = dataColumn.ColumnNumber();
                                                    var column = excelsheet.Column(dataColumnNumber);
                                                    if (dataColumnNumber > 1)
                                                    {
                                                        if (fileType == "s001") //反射率
                                                        {
                                                            if (siteCount == 1)
                                                            {
                                                                pointSite = "正";
                                                            }
                                                            else if (siteCount == 2)
                                                            {
                                                                pointSite = "反";
                                                            }
                                                            pointValue = row.Cell(dataColumnNumber).Value.ToString();

                                                            //點資料放LIST或MODEL，再一次INSERT
                                                            //QmdId, QcRecordFileId, LotNumber, BarcodeId, point, pointSite, pointValue, Axis, Cavity
                                                            //將值拋到DA，QcMeasurePointData
                                                            if (pointValue != "")
                                                            {
                                                                QcMeasurePointData PointDataSet = new QcMeasurePointData()
                                                                {
                                                                    //QmdId = QmdId,
                                                                    QcRecordFileId = QcRecordFileId,
                                                                    BarcodeId = Convert.ToInt32(BarcodeId),
                                                                    LotNumber = LotNumber,
                                                                    Point = point,
                                                                    PointValue = pointValue,
                                                                    Axis = Axis,
                                                                    Cavity = Cavity,
                                                                    PointSite = pointSite,
                                                                };

                                                                PointDataSetList.Add(PointDataSet);
                                                            }

                                                            siteCount += 1;
                                                            if ((dataColumnNumber - 1) % 2 == 0)
                                                            {
                                                                siteCount = 1;
                                                            }
                                                        }
                                                        else if (fileType == "s101") //穿透率
                                                        {
                                                            pointSite = row.Cell(dataColumnNumber).Value.ToString();
                                                            Cavity = row.Cell(dataColumnNumber).Value.ToString();
                                                        }

                                                    }
                                                }
                                            }
                                        }
                                        else if (CompanyId == 2) //中揚
                                        {

                                        }

                                    }
                                    #endregion
                                }
                                else if (filePath.IndexOf("csv") != -1)
                                {
                                    #region //資料解析 csv
                                    //if (CompanyId == 4) //晶彩
                                    {
                                        Encoding encoding = Encoding.UTF8;
                                        string line;
                                        using (StreamReader st = new StreamReader(filePath, encoding))
                                        {
                                            line = st.ReadLine();
                                            line = st.ReadLine();

                                            while ((line = st.ReadLine()) != null)
                                            {
                                                string[] lineData = line.Split(',');
                                                string pointSite = "";
                                                string pointValue = "";
                                                string Cavity = "";
                                                string Axis = "";
                                                int siteCount = 1;

                                                for (int i = 1; i < lineData.Length; i++)
                                                {
                                                    if (lineData[i] == "0")
                                                    {
                                                        break;
                                                    }
                                                    if (siteCount == 1)
                                                    {
                                                        pointSite = "正";
                                                    }
                                                    else if (siteCount == 2)
                                                    {
                                                        pointSite = "反";
                                                    }
                                                    string point = lineData[0];
                                                    pointValue = lineData[i];

                                                    if (pointValue != "")
                                                    {
                                                        QcMeasurePointData PointDataSet = new QcMeasurePointData()
                                                        {
                                                            //QmdId = QmdId,
                                                            QcRecordFileId = QcRecordFileId,
                                                            BarcodeId = Convert.ToInt32(BarcodeId),
                                                            LotNumber = LotNumber,
                                                            Point = point,
                                                            PointValue = pointValue,
                                                            Axis = Axis,
                                                            Cavity = Cavity,
                                                            PointSite = pointSite,
                                                        };

                                                        PointDataSetList.Add(PointDataSet);
                                                    }

                                                    siteCount += 1;
                                                    if (i % 2 == 0)
                                                    {
                                                        siteCount = 1;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    //else if (CompanyId == 2) //中揚
                                    {

                                    }
                                        #endregion
                                    }
                            }
                            #region //INSERT Point Data
                            dataRequest = qcDataManagementDA.AddQcMeasurePointData(PointDataSetList);
                            #endregion
                        }
                    }
                    else
                    {
                        throw new SystemException("查無檔案資料!!");
                    }
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                logger.Error(e.Message);
            }
        }

        #endregion
        #endregion

        #region //For SRM API
        #region //GetQcDeliveryInspectionForApi 取得供應商出貨檢驗單據資料(Api)
        [HttpPost]
        [Route("api/SCM/GetQcDeliveryInspection")]
        public void GetQcDeliveryInspectionForApi(string SecretKey = "", string CompanyNo = "", int QcDeliveryInspectionId = -1, string PoErpFullNo = "", int SupplierId = -1, int PoUserId = -1, string SearchKey = ""
            , string StartDate = "", string EndDate = "", int draw = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetQcDeliveryInspection");
                #endregion

                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQcDeliveryInspectionForApi(QcDeliveryInspectionId, PoErpFullNo, SupplierId, PoUserId, SearchKey
                    , StartDate, EndDate, CompanyNo, draw
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
        #endregion

        #region //FOR EIP API
        #region //GetQmmDetail 取得量測機台資料
        [HttpPost]
        [Route("api/CR/GetQmmDetail")]

        public void GetQmmDetailEIP(int QmmDetailId = -1, int ShopId = -1, string MachineNumber = "", int[] CustomerIds = null)
        {
            try
            {
                #region //Request
                qcDataManagementDA = new QcDataManagementDA();
                dataRequest = qcDataManagementDA.GetQmmDetailEIP(QmmDetailId, ShopId, MachineNumber, CustomerIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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