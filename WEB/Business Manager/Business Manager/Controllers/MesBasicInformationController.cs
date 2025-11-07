using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using MESDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic; 
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Xceed.Document.NET;

namespace Business_Manager.Controllers
{
    public class MesBasicInformationController : WebController
    {
        private MesBasicInformationDA mesBasicInformationDA = new MesBasicInformationDA();

        #region //View
        public ActionResult ProdModeManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult WorkShopManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MachineManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TimeIntervalManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DeviceManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProcessManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProdModeShiftManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProdUnitManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult WarehouseManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProcessParameterManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TrayManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UserEventSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MachineEventSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProcessEventSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProcessParameterQcItem()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult LensCarrierRingManagment()
        {
            //套環資料管理
            ViewLoginCheck();

            return View();
        }

        public ActionResult RingTransactionManagment()
        {
            //套環庫存異動管理
            ViewLoginCheck();

            return View();
        }

        #endregion

        #region //Get
        #region //GetCompanyProdMode 取得特定公司的生產模式
        [HttpPost]
        public void GetCompanyProdMode(int CompanyId=-1,int ModeId = -1, string ModeNo = "", string ModeName = "", string Status = "", string BarcodeCtrl = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetCompanyProdMode(CompanyId,ModeId, ModeNo, ModeName, Status, BarcodeCtrl
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

        #region //GetProdMode 取得生產模式
        [HttpPost]
        public void GetProdMode(int ModeId = -1, string ModeNo = "", string ModeName = "", string Status = "", string BarcodeCtrl="", string ScrapRegister =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProdMode(ModeId, ModeNo, ModeName, Status, BarcodeCtrl, ScrapRegister
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

        #region //GetProdModeShift 取得生產模式班次資料
        [HttpPost]
        public void GetProdModeShift(int ModeShiftId = -1, int ModeId = -1, int ShiftId = -1, string EffectiveStartDate ="", string EffectiveEndDate = ""
            , string ExpirationStartDate = "", string ExpirationEndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProdModeShift(ModeShiftId, ModeId, ShiftId, EffectiveStartDate, EffectiveEndDate
                    , ExpirationStartDate, ExpirationEndDate, Status
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

        #region //GetWorkShop 取得車間資料
        [HttpPost]
        public void GetWorkShop(int ShopId = -1, string ShopNo = "", string ShopName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WorkShopManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetWorkShop(ShopId, ShopNo, ShopName, Status
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

        #region //GetMachine 取得機台資料
        [HttpPost]
        public void GetMachine(int MachineId = -1, int ShopId = -1, string MachineNo = "", string MachineName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetMachine(MachineId, ShopId, MachineNo, MachineName, Status
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

        #region //GetWorkingMachine 取得機台資料
        [HttpPost]
        public void GetWorkingMachine(int MachineId = -1, int ShopId = -1, string MachineNo = "", string MachineName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "read");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetWorkingMachine(MachineId, ShopId, MachineNo, MachineName, Status
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

        #region //GetMachineAsset 取得機台資產編號
        [HttpPost]
        public void GetMachineAsset(int MachineId = -1, string AssetNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "asset");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetMachineAsset(MachineId, AssetNo
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

        #region //GetMachineLog 取得機台日誌
        [HttpPost]
        public void GetMachineLog(int MachineId = -1, string StartDate = "", string EndDate = "", string OperatingStatus = "", int UserId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "log");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetMachineLog(MachineId, StartDate, EndDate, OperatingStatus, UserId
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

        #region //GetTimeInterval 取得時間區段資料
        [HttpPost]
        public void GetTimeInterval(int IntervalId = -1, string BeginTime = "", string EndTime = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TimeIntervalManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetTimeInterval(IntervalId, BeginTime, EndTime
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

        #region //GetDevice 取得裝置資料
        [HttpPost]
        public void GetDevice(int DeviceId = -1, string DeviceType = "", string DeviceName = "", string DeviceIdentifierCode = "", string DeviceAuthority = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetDevice(DeviceId, DeviceType, DeviceName, DeviceIdentifierCode, DeviceAuthority, Status
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

        #region //GetDeviceMachine 取得裝置資料
        [HttpPost]
        public void GetDeviceMachine(int DeviceId = -1, int MachineId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetDeviceMachine(DeviceId, MachineId
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

        #region //GetProcess 取得製程資料
        [HttpPost]
        public void GetProcess(int ProcessId = -1, string ProcessNo = "", string ProcessName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcess(ProcessId, ProcessNo, ProcessName, Status
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

        #region //GetProdUnit 取得生產單元資料
        [HttpPost]
        public void GetProdUnit(int UnitId = -1, string UnitNo = "", string UnitName = "", string CheckStatus = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProdUnitManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProdUnit(UnitId, UnitNo, UnitName, CheckStatus, Status
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

        #region //GetWarehouse 取得庫房基本資料
        [HttpPost]
        public void GetWarehouse(int WarehouseId = -1, int ShopId = -1, string WarehouseNo = "", string WarehouseName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetWarehouse(WarehouseId, ShopId, WarehouseNo, WarehouseName, Status
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

        #region //GetWarehouseLocation 取得庫房儲位基本資料
        [HttpPost]
        public void GetWarehouseLocation(int WarehouseId = -1, int LocationId = -1, string LocationNo = "", string LocationName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "location");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetWarehouseLocation(WarehouseId, LocationId, LocationNo, LocationName, Status
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

        #region //GetProcessParameter 取得製程參數資料
        [HttpPost]
        public void GetProcessParameter(int ParameterId = -1, int ProcessId = -1, int DepartmentId = -1, int ModeId = -1, string ProcessCheckStatus = ""
            , string PreCollectionStatus = "", string PostCollectionStatus = "", string NgToBarcode = "", string PassingMode = "", string ProcessCheckType = ""
            , string Status = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcessParameter(ParameterId, ProcessId, DepartmentId, ModeId, ProcessCheckStatus, PreCollectionStatus
                    , PostCollectionStatus, NgToBarcode, PassingMode, ProcessCheckType
                    , Status, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessMachine 取得製程機台資料
        [HttpPost]
        public void GetProcessMachine(int ParameterId = -1, int MachineId = -1, string BatchStatus = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcessMachine(ParameterId, MachineId, BatchStatus, Status
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

        #region //GetProcessProductionUnit 取得製程生產單元資料
        [HttpPost]
        public void GetProcessProductionUnit(int ParameterId = -1, int UnitId = -1, int SortNumber = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "unit");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcessProductionUnit(ParameterId, UnitId, SortNumber, Status
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

        #region //GetTray 取得托盤資料
        [HttpPost]
        public void GetTray(int TrayId = -1, string MoNo = "", string BarcodeNo = "", string TrayNo = "", string TrayName = "", string TrayType = "", string BindStatus = "", string Status = "", string DeleteBatch = ""
            , string CreateUserId = "", string CreatTimeStart = "", string CreatTimeEnd = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetTray(TrayId, MoNo, BarcodeNo, TrayNo, TrayName, TrayType, BindStatus, Status, DeleteBatch
                    , CreateUserId, CreatTimeStart, CreatTimeEnd, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTrayBarcodeLog 取得托盤曾經綁定條碼資料
        [HttpPost]
        public void GetTrayBarcodeLog(int TrayId = -1, string TrayNo = "", string TrayNoList = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "read");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetTrayBarcodeLog(TrayId, TrayNo, TrayNoList
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

        #region //GetKeyence 取得Keyence資料
        [HttpPost]
        public void GetKeyence(int KeyenceId = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetKeyence(KeyenceId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserEventSetting 取得人員事件設定
        [HttpPost]
        public void GetUserEventSetting(int UserEventItemId = -1, string UserEventItemNo = "", string UserEventItemName = "", string OrderBy="", int PageIndex=-1, int PageSize=-1)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetUserEventSetting(UserEventItemId, UserEventItemNo, UserEventItemName, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMachineEventSetting 取得機台事件設定
        [HttpPost]
        public void GetMachineEventSetting(int ShopId=-1, int MachineId=-1, string ShopName="", string MachineName="", int MachineEventItemId = -1, string MachineEventNo = "", string MachineEventName = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetMachineEventSetting(ShopId, MachineId, ShopName, MachineName, MachineEventItemId, MachineEventNo, MachineEventName, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessEventSetting 取得加工事件設定
        [HttpPost]
        public void GetProcessEventSetting(int ProcessEventItemId=-1, int ProcessId=-1, string TypeNo="", string ProcessEventName="", int ParameterId = -1, int ModeId=-1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcessEventSetting(ProcessEventItemId, ProcessId, TypeNo, ProcessEventName, ParameterId, ModeId, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessParameterQcItem //取得製程參數量測項目
        [HttpPost]
        public void GetProcessParameterQcItem(int ParameterId = -1)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcessParameterQcItem(ParameterId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRoutingItemQcItem //途程量測項目 
        [HttpPost]
        public void GetRoutingItemQcItem(int RoutingItemId = -1 , int ItemProcessId = -1)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetRoutingItemQcItem(RoutingItemId, ItemProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLensCarrierRingManagment 取得套環基本資料
        [HttpPost]
        public void GetLensCarrierRingManagment(int RingId = -1, string ModelName = "", string RingCode = "", string Customer = "", string Status = "", string SelStockAlert = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetLensCarrierRingManagment(RingId, ModelName, RingCode, Customer, Status, SelStockAlert, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRingTransactionManagment 套環庫存異動管理
        [HttpPost]
        public void GetRingTransactionManagment( int RingTransId = -1, string ModelName = "", string RingName = "", string Status = ""
            , string StartTransDate = "", string EndTransDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "read,constrained-data");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetRingTransactionManagment(RingTransId, ModelName, RingName, StartTransDate, EndTransDate, Status, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddProdMode 生產模式新增
        [HttpPost]
        public void AddProdMode(int ModeId = -1, string ModeNo = "", string ModeName = "", string ModeDesc = "", string BarcodeCtrl ="", string ScrapRegister = ""
            , string Source = "", string PVTQCFlag = "", string NgToBarcode = "", string TrayBarcode = "", string LotStatus = "", string OutputBarcodeFlag = "", string MrType = ""
            , string OQcCheckType = "")
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProdMode(ModeId, ModeNo, ModeName, ModeDesc, BarcodeCtrl, ScrapRegister,
                    Source, PVTQCFlag, NgToBarcode, TrayBarcode, LotStatus, OutputBarcodeFlag, MrType, OQcCheckType
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

        #region //AddProdModeShift 生產模式班次資料新增
        [HttpPost]
        public void AddProdModeShift(int ModeId = -1, int ShiftId = -1, string EffectiveDate = "", string ExpirationDate = "")
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProdModeShift(ModeId, ShiftId, EffectiveDate, ExpirationDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddWorkShop 車間資料新增
        [HttpPost]
        public void AddWorkShop(string ShopNo = "", string ShopName = "", string ShopDesc = "", string Location = "", string Floor = "")
        {
            try
            {
                WebApiLoginCheck("WorkShopManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddWorkShop(ShopNo, ShopName, ShopDesc, Location, Floor);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMachine 機台資料新增
        [HttpPost]
        public void AddMachine(int ShopId = -1, string MachineNo = "", string MachineName = "", string MachineDesc = "")
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddMachine(ShopId, MachineNo, MachineName, MachineDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMachineAsset 機台資產編號資料新增
        [HttpPost]
        public void AddMachineAsset(int MachineId = -1, string AssetNo = "")
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "asset");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddMachineAsset(MachineId, AssetNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTimeInterval 時間區段新增
        [HttpPost]
        public void AddTimeInterval(string BeginTime = "", string EndTime = "")
        {
            try
            {
                WebApiLoginCheck("TimeIntervalManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddTimeInterval(BeginTime, EndTime);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDevice 裝置資料新增
        [HttpPost]
        public void AddDevice(string DeviceType = "", string DeviceName = "", string DeviceDesc = "", string DeviceIdentifierCode = "", string DeviceAuthority = "")
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddDevice(DeviceType, DeviceName, DeviceDesc, DeviceIdentifierCode, DeviceAuthority);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDeviceMachine 裝置機台綁定
        [HttpPost]
        public void AddDeviceMachine(int DeviceId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddDeviceMachine(DeviceId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProcess 製程資料新增
        [HttpPost]
        public void AddProcess(string ProcessNo = "", string ProcessName = "", string ProcessDesc = "")
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProcess(ProcessNo, ProcessName, ProcessDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProdUnit 生產單元資料新增
        [HttpPost]
        public void AddProdUnit(string UnitNo = "", string UnitName = "", string UnitDesc = "", string CheckStatus = "")
        {
            try
            {
                WebApiLoginCheck("ProdUnitManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProdUnit(UnitNo, UnitName, UnitDesc, CheckStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddWarehouse 庫房基本資料新增
        [HttpPost]
        public void AddWarehouse(int ShopId = -1, string WarehouseNo = "", string WarehouseName = "", string WarehouseDesc = "")
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddWarehouse(ShopId, WarehouseNo, WarehouseName, WarehouseDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddWarehouseLocation 庫房基本資料新增
        [HttpPost]
        public void AddWarehouseLocation(int WarehouseId = -1, string LocationNo = "", string LocationName = "", string LocationDesc = "")
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "location");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddWarehouseLocation(WarehouseId, LocationNo, LocationName, LocationDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProcessParameter 製程參數資料新增
        [HttpPost]
        public void AddProcessParameter(int ProcessId = -1, int ModeId = -1, int DepartmentId = -1, string ProcessCheckStatus = "", string PreCollectionStatus = "", string PackageFlag = ""
            , string PostCollectionStatus = "", string NgToBarcode = "", string PassingMode = "", string ProcessCheckType = "", string ConsumeFlag = "")
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProcessParameter(ProcessId, ModeId, DepartmentId, ProcessCheckStatus, PreCollectionStatus, PackageFlag
                    , PostCollectionStatus, NgToBarcode, PassingMode, ProcessCheckType, ConsumeFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProcessMachine 製程機台資料新增
        [HttpPost]
        public void AddProcessMachine(int ParameterId = -1, int MachineId = -1, int ToolCount = -1, string KeyenceFlag = "", int KeyenceId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProcessMachine(ParameterId, MachineId, ToolCount, KeyenceFlag, KeyenceId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProcessProductionUnit 製程生產單元新增
        [HttpPost]
        public void AddProcessProductionUnit(int ParameterId = -1, int UnitId = -1, int SortNumber = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "unit");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProcessProductionUnit(ParameterId, UnitId, SortNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTray 托盤資料新增
        [HttpPost]
        public void AddTray(string TrayPrefix = "", string TrayName = "", int TrayCapacity = -1, int SerialNumber = -1, int Fabrication = -1, int Serial = -1, string SuffixCode =""
            , string Remark = "", int ViewCompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddTray(TrayPrefix, TrayName, TrayCapacity, SerialNumber, Fabrication, Serial, SuffixCode
                    , Remark, ViewCompanyId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddUserEventSetting 人員事件設定新增
        [HttpPost]
        public void AddUserEventSetting(string UserEventItemName="")
        {
            try
            {
                WebApiLoginCheck("UserEventSetting", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddUserEventSetting(UserEventItemName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMachineEventSetting 機台事件設定新增
        [HttpPost]
        public void AddMachineEventSetting(int MachineId=-1, string MachineEventNo="", string MachineEventName="")
        {
            try
            {
                WebApiLoginCheck("MachineEventSetting", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddMachineEventSetting(MachineId, MachineEventNo, MachineEventName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProcessEventSetting 加工事件設定新增
        [HttpPost]
        public void AddProcessEventSetting(int ParameterId=-1, string ProcessEventNo="", string ProcessEventName="", string ProcessEventType="")
        {
            try
            {
                WebApiLoginCheck("ProcessEventSetting", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddProcessEventSetting(ParameterId, ProcessEventNo, ProcessEventName, ProcessEventType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDefaultRoutingQcitem 途程品號帶入製程量測項目
        [HttpPost]
        public void AddDefaultRoutingQcitem(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddDefaultRoutingQcitem(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDefaultRoutingQcitem2 途程品號帶入製程量測項目(Excel帶入)
        [HttpPost]
        public void AddDefaultRoutingQcitem2(int RoutingItemId = -1, int ItemProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddDefaultRoutingQcitem2(RoutingItemId, ItemProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDefaultRoutingQcitem3 途程品號帶入製程量測項目(確認製程時 全部途程品號新增)
        [HttpPost]
        public void AddDefaultRoutingQcitem3(int RoutingId = -1)
        {
            try
            {
                WebApiLoginCheck("RoutingManagement", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddDefaultRoutingQcitem3(RoutingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddLensCarrierRing 套環資料新增
        [HttpPost]
        public void AddLensCarrierRing(string ModelName = "", string RingName = "", string Remarks = "", string HoleCount = ""
            , string RingSpec = "", string RingCode = "", string RingShape = "", string Customer = "", decimal DailyDemand = -1, decimal SafetyStock = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddLensCarrierRing(ModelName, RingName, Remarks, HoleCount, RingSpec, RingCode, RingShape, Customer, DailyDemand, SafetyStock);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRingTransaction 套環庫存異動新增
        [HttpPost]
        public void AddRingTransaction(int RingId = -1, string ModelName = "", string RingName = "", string TransType = ""
            , string TransDate = "", int Quantity = -1)
        {
            try
            {
                WebApiLoginCheck("RingTransactionManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddLensCarrierRing(RingId, ModelName, RingName, TransType, TransDate, Quantity);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateProdMode 生產模式更新
        [HttpPost]
        public void UpdateProdMode(int ModeId = -1, string ModeName = "", string ModeDesc = "", string BarcodeCtrl = "", string ScrapRegister = ""
            , string Source = "", string PVTQCFlag = "", string NgToBarcode = "", string TrayBarcode = "", string LotStatus = "", string OutputBarcodeFlag = "", string MrType = ""
            , string OQcCheckType ="")
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProdMode(ModeId, ModeName, ModeDesc, BarcodeCtrl, ScrapRegister,
                    Source, PVTQCFlag, NgToBarcode, TrayBarcode, LotStatus, OutputBarcodeFlag, MrType,
                    OQcCheckType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdModeStatus 生產模式狀態更新
        [HttpPost]
        public void UpdateProdModeStatus(int ModeId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProdModeStatus(ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdModeShift 生產模式班次資料更新
        [HttpPost]
        public void UpdateProdModeShift(int ModeShiftId = -1, int ShiftId = -1, string EffectiveDate = "", string ExpirationDate = "")
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProdModeShift(ModeShiftId, ShiftId, EffectiveDate, ExpirationDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdModeShiftStatus 生產模式班次狀態更新
        [HttpPost]
        public void UpdateProdModeShiftStatus(int ModeShiftId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProdModeShiftStatus(ModeShiftId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWorkShop 車間資料更新
        [HttpPost]
        public void UpdateWorkShop(int ShopId = -1, string ShopName = "", string ShopDesc = "", string Location = "", string Floor = "")
        {
            try
            {
                WebApiLoginCheck("WorkShopManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateWorkShop(ShopId, ShopName, ShopDesc, Location, Floor);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWorkShopStatus 車間狀態更新
        [HttpPost]
        public void UpdateWorkShopStatus(int ShopId = -1)
        {
            try
            {
                WebApiLoginCheck("WorkShopManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateWorkShopStatus(ShopId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMachine 機台資料更新
        [HttpPost]
        public void UpdateMachine(int MachineId = -1, int ShopId = -1, string MachineName = "", string MachineDesc = "")
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateMachine(MachineId, ShopId, MachineName, MachineDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMachineStatus 機台狀態更新
        [HttpPost]
        public void UpdateMachineStatus(int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateMachineStatus(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDevice 裝置資料更新
        [HttpPost]
        public void UpdateDevice(int DeviceId = -1, string DeviceType = "", string DeviceName = "", string DeviceDesc = "", string DeviceIdentifierCode = "", string DeviceAuthority = "")
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateDevice(DeviceId, DeviceType, DeviceName, DeviceDesc, DeviceIdentifierCode, DeviceAuthority);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDeviceIdStatus 裝置狀態更新
        [HttpPost]
        public void UpdateDeviceIdStatus(int DeviceId = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateDeviceIdStatus(DeviceId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDeviceMachine 裝置機台綁定資料更新
        [HttpPost]
        public void UpdateDeviceMachine(int DeviceId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateDeviceMachine(DeviceId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcess 車間資料更新
        [HttpPost]
        public void UpdateProcess(int ProcessId = -1, string ProcessName = "", string ProcessDesc = "")
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcess(ProcessId, ProcessName, ProcessDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessStatus 製程狀態更新
        [HttpPost]
        public void UpdateProcessStatus(int ProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessStatus(ProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdUnit 生產單元資料
        [HttpPost]
        public void UpdateProdUnit(int UnitId = -1, string UnitName = "", string UnitDesc = "", string CheckStatus = "")
        {
            try
            {
                WebApiLoginCheck("ProdUnitManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProdUnit(UnitId, UnitName, UnitDesc, CheckStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdUnitStatus 生產單元狀態更新
        [HttpPost]
        public void UpdateProdUnitStatus(int UnitId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdUnitManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProdUnitStatus(UnitId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWarehouse 庫房基本資料更新
        [HttpPost]
        public void UpdateWarehouse(int WarehouseId = -1, int ShopId = -1, string WarehouseName = "", string WarehouseDesc = "")
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateWarehouse(WarehouseId, ShopId, WarehouseName, WarehouseDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWarehouseStatus 庫房基本資料狀態更新
        [HttpPost]
        public void UpdateWarehouseStatus(int WarehouseId = -1)
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateWarehouseStatus(WarehouseId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWarehouseLocation 庫房儲位基本資料更新
        [HttpPost]
        public void UpdateWarehouseLocation(int WarehouseId = -1, int LocationId = -1, string LocationName = "", string LocationDesc = "")
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "location");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateWarehouseLocation(WarehouseId, LocationId, LocationName, LocationDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateWarehouseStatus 庫房儲位基本資料狀態更新
        [HttpPost]
        public void UpdateWarehouseLocationStatus(int LocationId = -1)
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "location");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateWarehouseLocationStatus(LocationId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessParameter 製程參數資料更新
        [HttpPost]
        public void UpdateProcessParameter(int ParameterId = -1, int ProcessId = -1, int ModeId = -1, int DepartmentId = -1, string ProcessCheckStatus = "", string PreCollectionStatus = ""
            , string PostCollectionStatus = "", string NgToBarcode = "", string PassingMode = "", string ProcessCheckType = "", string ConsumeFlag = "")
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessParameter(ParameterId, ProcessId, ModeId, DepartmentId, ProcessCheckStatus, PreCollectionStatus
                    , PostCollectionStatus, NgToBarcode, PassingMode, ProcessCheckType, ConsumeFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessParameterStatus 製程參數資料狀態更新
        [HttpPost]
        public void UpdateProcessParameterStatus(int ParameterId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessParameterStatus(ParameterId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBatchStatus 製程機台資料是否支援編成
        [HttpPost]
        public void UpdateBatchStatus(int ParameterId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "unit");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateBatchStatus(ParameterId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        
        #region //UpdateProcessMachine 更新製程機台資料
        [HttpPost]
        public void UpdateProcessMachine(int ParameterId = -1, int MachineId = -1, int ToolCount = -1, string KeyenceFlag = "", int KeyenceId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessMachine(ParameterId, MachineId, ToolCount, KeyenceFlag, KeyenceId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateParameterMachineStatus 製程機台資料狀態更新
        [HttpPost]
        public void UpdateParameterMachineStatus(int ParameterId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateParameterMachineStatus(ParameterId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateKeyenceFlag 更新製程機台是否支援Keyence設備
        [HttpPost]
        public void UpdateKeyenceFlag(int ParameterId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateKeyenceFlag(ParameterId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessProductionUnit 製程生產單元資料更新
        [HttpPost]
        public void UpdateProcessProductionUnit(int ParameterId = -1,  int UnitId = -1, int SortNumber = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "unit");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessProductionUnit(ParameterId, UnitId, SortNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessProductionUnitStatus 製程生產單元狀態更新
        [HttpPost]
        public void UpdateProcessProductionUnitStatus(int ParameterId = -1, int UnitId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "unit");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessProductionUnitStatus(ParameterId, UnitId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTimeInterval 時間區段更新
        [HttpPost]
        public void UpdateTimeInterval(int IntervalId = -1, string BeginTime = "", string EndTime = "")
        {
            try
            {
                WebApiLoginCheck("TimeIntervalManagment", "add");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateTimeInterval(IntervalId,BeginTime, EndTime);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTray 托盤資料更新
        [HttpPost]
        public void UpdateTray(int TrayId = -1, string TrayName = "", int TrayCapacity = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateTray(TrayId, TrayName, TrayCapacity, Remark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTrayStatus 托盤狀態更新
        [HttpPost]
        public void UpdateTrayStatus(int TrayId = -1)
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "status-switch");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateDeviceIdStatus(TrayId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTrayUnBind 托盤解除綁定
        [HttpPost]
        public void UpdateTrayUnBind(int TrayId = -1, string TrayNo = "", string Place = "")
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateTrayUnBind(TrayId, TrayNo, Place);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        
        #region //UpdateUserEventSetting 人員事件設定更新
        [HttpPost]
        public void UpdateUserEventSetting(int UserEventItemId = -1, string UserEventItemName = "")
        {
            try
            {
                WebApiLoginCheck("UserEventSetting", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateUserEventSetting(UserEventItemId, UserEventItemName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMachineEventSetting 機台事件設定更新
        [HttpPost]
        public void UpdateMachineEventSetting(int MachineId=-1, int MachineEventItemId = -1, string MachineEventName = "")
        {
            try
            {
                WebApiLoginCheck("MachineEventSetting", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateMachineEventSetting(MachineId, MachineEventItemId, MachineEventName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessEventSetting 加工事件設定更新
        [HttpPost]
        public void UpdateProcessEventSetting(int ProcessEventItemId=-1, string ProcessEventName="")
        {
            try
            {
                WebApiLoginCheck("ProcessEventSetting", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessEventSetting(ProcessEventItemId, ProcessEventName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProcessParameterQcitemExcel 上傳製程量測參數Excel資料
        [HttpPost]
        public void UpdateProcessParameterQcitemExcel(int ParameterId = -1, string ExcelJson = "")
        {
            try
            {
                //WebApiLoginCheck("ProcessEventSetting", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateProcessParameterQcitemExcel(ParameterId, ExcelJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoutingItemQcItemExcel 上傳途程量測Excel資料
        [HttpPost]
        public void UpdateRoutingItemQcItemExcel(int RoutingItemId = -1, string ExcelJson = "", int ItemProcessId = -1)
        {
            try
            {
                //WebApiLoginCheck("ProcessEventSetting", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateRoutingItemQcItemExcel(RoutingItemId, ExcelJson, ItemProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLensCarrierRing 套環資料更新
        [HttpPost]
        public void UpdateLensCarrierRing(int RingId = -1, string ModelName = "", string RingName = "", string Remarks = "", string HoleCount = ""
            , string RingSpec = "", string RingCode = "", string RingShape = "", string Customer = "", decimal DailyDemand = -1, decimal SafetyStock = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateLensCarrierRing(RingId, ModelName, RingName, Remarks, HoleCount, RingSpec, RingCode, RingShape, Customer, DailyDemand, SafetyStock);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLensCarrierRingStatus 套環資料状态更新
        [HttpPost]
        public void UpdateLensCarrierRingStatus(int RingId = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "void");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateLensCarrierRingStatus(RingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion 

        #region //ImportExcelRing 解析Excel新增套環資料
        [HttpPost]
        public void ImportExcelRing(string FileId = "")
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "add");

                if (FileId.Length <= 0) throw new SystemException("必須上傳檔案!!");

                List<LCRExcelFormat> lcrExcelFormats = new List<LCRExcelFormat>();

                string[] FileIds = FileId.Split(',');
                foreach (var fileId in FileIds)
                {
                    #region //取得File資料
                    LCRFileInfo fileInfo = mesBasicInformationDA.GetLCRFileInfoById(Convert.ToInt32(fileId));
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
                                    LCRExcelFormat lcrExcelFormat = new LCRExcelFormat()
                                    {
                                        ModelName = table.Cell(i, 1).Value.ToString(),
                                        RingName = table.Cell(i, 2).Value.ToString(),
                                        Remarks = table.Cell(i, 3).Value.ToString(),
                                        HoleCount = int.TryParse(table.Cell(i, 4).Value.ToString(), out int holeCount) ? holeCount : 0,
                                        RingSpec = table.Cell(i, 5).Value.ToString(),
                                        RingCode = table.Cell(i, 6).Value.ToString(),
                                        RingShape = table.Cell(i, 7).Value.ToString(),
                                        Customer = table.Cell(i, 8).Value.ToString(),
                                        DailyDemand = decimal.Parse(table.Cell(i, 9).Value.ToString()),
                                        SafetyStock = decimal.Parse(table.Cell(i, 10).Value.ToString()),
                                    };
                                    lcrExcelFormats.Add(lcrExcelFormat);
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
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddLensCarrierRing(lcrExcelFormats);
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

        #region //UpdateRingTransaction 套環庫存異動更新
        [HttpPost]
        public void UpdateRingTransaction(int RingTransId = -1, int RingId = -1, string TransType = "", int Quantity = -1)
        {
            try
            {
                WebApiLoginCheck("RingTransactionManagment", "update");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateRingTransaction(RingTransId, RingId, TransType, Quantity);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLensCarrierRingStatus 套環庫存異動状态更新
        [HttpPost]
        public void UpdateRingTransactionStatus(int RingTransId = -1)
        {
            try
            {
                WebApiLoginCheck("RingTransactionManagment", "void");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.UpdateRingTransactionStatus(RingTransId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion 

        #region //ImportExcelRingTrans 解析Excel新增套環庫存異動資料
        [HttpPost]
        public void ImportExcelRingTrans(string FileId = "")
        {
            try
            {
                WebApiLoginCheck("RingTransactionManagment", "add");

                if (FileId.Length <= 0) throw new SystemException("必須上傳檔案!!");

                List<LCRTransExcelFormat> lcrTransExcelFormats = new List<LCRTransExcelFormat>();

                string[] FileIds = FileId.Split(',');
                foreach (var fileId in FileIds)
                {
                    #region //取得File資料
                    LCRFileInfo fileInfo = mesBasicInformationDA.GetLCRFileInfoById(Convert.ToInt32(fileId));
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
                                    

                                    LCRTransExcelFormat lcrTransExcelFormat = new LCRTransExcelFormat()
                                    {
                                        ModelName = table.Cell(i, 1).Value.ToString(),
                                        TransType = table.Cell(i, 2).Value.ToString(),
                                        TransDate = table.Cell(i, 3).Value.ToString(),
                                        Quantity = int.TryParse(table.Cell(i, 4).Value.ToString(), out int quantity) ? quantity : 0,
                                    };
                                    lcrTransExcelFormats.Add(lcrTransExcelFormat);
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
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.AddRingTransaction(lcrTransExcelFormats);
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
         
        #endregion

        #region //Delete
        #region //DeleteProdModeShift 生產模式班次資料刪除
        [HttpPost]
        public void DeleteProdModeShift(int ModeShiftId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteProdModeShift(ModeShiftId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMachineAsset 機台資產編號刪除
        [HttpPost]
        public void DeleteMachineAsset(int MachineId = -1, string AssetNo = "")
        {
            try
            {
                WebApiLoginCheck("MachineManagment", "asset");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteMachineAsset(MachineId, AssetNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDevice 裝置刪除
        [HttpPost]
        public void DeleteDevice(int DeviceId = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteDevice(DeviceId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDeviceMachine 裝置機台綁定刪除
        [HttpPost]
        public void DeleteDeviceMachine(int DeviceId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("DeviceManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteDeviceMachine(DeviceId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProcess 製程刪除
        [HttpPost]
        public void DeleteProcess(int ProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteProcess(ProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProdUnit 生產單元資料刪除
        [HttpPost]
        public void DeleteProdUnit(int UnitId = -1)
        {
            try
            {
                WebApiLoginCheck("ProdUnitManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteProdUnit(UnitId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteWarehouseLocation 裝置機台綁定刪除
        [HttpPost]
        public void DeleteWarehouseLocation(int WarehouseId = -1, int LocationId = -1)
        {
            try
            {
                WebApiLoginCheck("WarehouseManagment", "location");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteWarehouseLocation(WarehouseId, LocationId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProcessMachine 製程機台資料刪除
        [HttpPost]
        public void DeleteProcessMachine(int ParameterId = -1, int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "machine");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteProcessMachine(ParameterId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProcessProductionUnit 製程生產單元刪除
        [HttpPost]
        public void DeleteProcessProductionUnit(int ParameterId = -1, int UnitId = -1)
        {
            try
            {
                WebApiLoginCheck("ProcessParameterManagment", "unit");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteProcessProductionUnit(ParameterId, UnitId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTray 托盤刪除
        [HttpPost]
        public void DeleteTray(int TrayId = -1)
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteTray(TrayId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTrayBatch 托盤刪除(批量)
        [HttpPost]
        public void DeleteTrayBatch(string TrayListY = "")
        {
            try
            {
                WebApiLoginCheck("TrayManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteTrayBatch(TrayListY);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteUserEventSetting 人員事件資料刪除
        [HttpPost]
        public void DeleteUserEventSetting(int UserEventItemId = -1)
        {
            try
            {
                WebApiLoginCheck("UserEventSetting", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteUserEventSetting(UserEventItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMachineEventSetting 機台事件資料刪除
        [HttpPost]
        public void DeleteMachineEventSetting(int MachineEventItemId = -1)
        {
            try
            {
                //WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteMachineEventSetting(MachineEventItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProcessEventSetting 加工事件資料刪除
        [HttpPost]
        public void DeleteProcessEventSetting(int ProcessEventItemId = -1)
        {
            try
            {
                //WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteProcessEventSetting(ProcessEventItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteAllProcessParameterQcitem 刪除全部製程量測項目
        [HttpPost]
        public void DeleteAllProcessParameterQcitem(int ParameterId = -1)
        {
            try
            {
                //WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteAllProcessParameterQcitem(ParameterId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteAllRoutingItemQcItem 刪除全部途程量測項目
        [HttpPost]
        public void DeleteAllRoutingItemQcItem(int RoutingItemId = -1, int ItemProcessId = -1)
        {
            try
            {
                //WebApiLoginCheck("ProdModeManagment", "prod-mode-shift");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteAllRoutingItemQcItem(RoutingItemId, ItemProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteLensCarrierRing 套環資料刪除
        [HttpPost]
        public void DeleteLensCarrierRing(int RingId = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteLensCarrierRing(RingId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteRingTransaction 套環庫存異動資料刪除
        [HttpPost]
        public void DeleteRingTransaction(int RingTransId = -1)
        {
            try
            {
                WebApiLoginCheck("RingTransactionManagment", "delete");

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.DeleteRingTransaction(RingTransId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Excel
        #region //ExcelLensCarrierRingManagment 套環資料管理
        public void ExcelLensCarrierRingManagment(int RingId = -1, string ModelName = "", string RingCode = "", string Customer = "", string Status = "", string SelStockAlert = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int RedirectType = -1)
        {
            try
            {
                WebApiLoginCheck("LensCarrierRingManagment", "read,excel");
                List<string> OsrIdList = new List<string>();
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetLensCarrierRingManagment(RingId, ModelName, RingCode, Customer, Status, SelStockAlert, OrderBy, PageIndex, PageSize);
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
                    string excelFileName = "【MES2.0】套環資料管理";
                    string excelsheetName = "頁1";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] header = new string[] { "機種名", "套環命名", "備註", "孔數", "日需求量", "安全庫存", "當前庫存"
                        ,"庫存預警","套環規格","套環編碼","套環形狀","客戶"
                         };

                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 12;
                        worksheet.Style = defaultStyle;


                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        #region //BODY

                        foreach (var item in data)
                        {
                            var rowData = new[]
                            {
                                item.ModelName?.ToString() ?? "",
                                item.RingName?.ToString() ?? "",
                                item.Remarks?.ToString() ?? "",
                                item.HoleCount?.ToString() ?? "",
                                item.DailyDemand?.ToString() ?? "",
                                item.SafetyStock?.ToString() ?? "",
                                item.Quantity?.ToString() ?? "",
                                item.SafetyStockLevel?.ToString() ?? "",
                                item.RingSpec?.ToString() ?? "",
                                item.RingCode?.ToString() ?? "",
                                item.RingShape?.ToString() ?? "",
                                item.Customer?.ToString() ?? ""
                            };

                            rowIndex++;

                            for (int i = 0; i < rowData.Length; i++)
                            {
                                worksheet.Cell(BaseHelper.MergeNumberToChar(i + 1, rowIndex)).Value = rowData[i];
                            }
                        }
                        #endregion

                        #region //設定

                        #region //自適應欄寬
                        worksheet.Columns().AdjustToContents();
                        //worksheet.Column(1).Width = 14;
                        //worksheet.Column(2).Width = 16;

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
                        worksheet.SheetView.FreezeRows(1);
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

        #endregion

        #region //FOR EIP API
        #region //GetProdModeEIP 取得生產模式
        [HttpPost]
        [Route("api/CR/GetProdMode")]

        public void GetProdModeEIP(int ModeId = -1, string ModeNo = "", string ModeName = "", string Status = "", string BarcodeCtrl = "", string ScrapRegister = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProdModeEIP(ModeId, ModeNo, ModeName, Status, BarcodeCtrl, ScrapRegister
                    , OrderBy, PageIndex, PageSize, CustomerIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessEIP -- 取得製程資料
        [HttpPost]
        [Route("api/CR/GetProcess")]

        public void GetProcessEIP(int ProcessId = -1, string ProcessNo = "", string ProcessName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetProcessEIP(ProcessId, ProcessNo, ProcessName, Status, OrderBy, PageIndex
                    , PageSize, CustomerIds);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//API

        //#region//晶彩-套環雷刻機台【API01】取得雷刻套環編號
        //[HttpPost]
        //[Route("api/MES/apiGetLaserEngravedRingCode")]
        //public void GetLaserEngravedRingCode(string Company = "", string SecretKey = "",string UserNo=""
        //    , string TrayPrefix = "", string TrayName = "", int TrayCapacity = -1, int SerialNumber = -1, int Fabrication = -1, int Serial = -1, string SuffixCode = ""
        //    , string Remark = "")
        //{
        //    try
        //    {

        //        Fabrication = 1;
        //        SerialNumber = 0;               
        //        SuffixCode = "N";

        //        #region //Api金鑰驗證
        //        ApiKeyVerify(Company, SecretKey, "GetLaserEngravedRingCode");
        //        #endregion

        //        #region //Request
        //        mesBasicInformationDA = new MesBasicInformationDA();
        //        dataRequest = mesBasicInformationDA.apiAddTray(Company,UserNo,TrayPrefix, TrayName, TrayCapacity, SerialNumber, Fabrication, Serial, SuffixCode, Remark);
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
        //#endregion

        //#region//晶彩 套環雷刻機台 【API02】紀錄此次雷刻條碼
        //[HttpPost]
        //[Route("api/MES/apiUpdateLaserEngravedRingCode")]
        //public void UpdateLaserEngravedRingCode(string Company = "", string SecretKey = "", string UserNo=""
        //    , string TrayNo = "")
        //{
        //    try
        //    {
        //        #region //Api金鑰驗證
        //        ApiKeyVerify(Company, SecretKey, "UpdateLaserEngravedRingCode");
        //        #endregion

        //        #region //Request
        //        mesBasicInformationDA = new MesBasicInformationDA();
        //        dataRequest = mesBasicInformationDA.UpdateLaserEngravedRingCode(Company, UserNo, TrayNo);
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
        //#endregion

        //#region//晶彩 套環雷刻機台 【API03】取得雷射刻印人員
        //[HttpPost]
        //[Route("api/MES/apiGetLaserEngravedUser")]
        //public void GetLaserEngravedUser(string Company = "", string SecretKey = "", string LoginNo = "", string PassWord = "")
        //{
        //    try
        //    {
        //        #region //Api金鑰驗證
        //        ApiKeyVerify(Company, SecretKey, "GetLaserEngravedUser");
        //        #endregion

        //        #region //Request
        //        mesBasicInformationDA = new MesBasicInformationDA();
        //        dataRequest = mesBasicInformationDA.GetLaserEngravedUser(Company, LoginNo, PassWord);
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
        //#endregion

        #region//AI建議的  晶彩-套環雷刻機台【API01】取得雷刻套環編號
        [HttpPost]
        [Route("api/MES/apiGetLaserEngravedRingCode")]
        public async Task GetLaserEngravedRingCodeAsync(string Company = "", string SecretKey = "", string UserNo = ""
, string TrayPrefix = "", string TrayName = "", int TrayCapacity = -1, int Serial = -1, string Remark = "")
        {
            try
            {
                // 預設固定值
                int SerialNumber = 0;
                int Fabrication = 1;
                string SuffixCode = "N";

                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetLaserEngravedRingCode");
                #endregion

                #region //Request with Retry - 直接使用apiAddTrayAsync的回傳
                mesBasicInformationDA = new MesBasicInformationDA();

                // 記錄傳入參數用於除錯
                logger.Info($"API參數 - Company: {Company}, UserNo: {UserNo}, TrayPrefix: {TrayPrefix}, TrayName: {TrayName}, TrayCapacity: {TrayCapacity}, Serial: {Serial}");

                // 直接呼叫 apiAddTrayAsync，它已經回傳正確格式的JSON
                dataRequest = await RetryAsync(async () =>
                    await mesBasicInformationDA.apiAddTrayAsync(Company, UserNo, TrayPrefix, TrayName, TrayCapacity,
                        SerialNumber, Fabrication, Serial, SuffixCode, Remark),
                    maxRetries: 3);
                #endregion

                #region //Response - apiAddTrayAsync已經回傳正確格式，直接解析即可
                try
                {
                    var response = JObject.Parse(dataRequest);
                    logger.Info($"apiAddTrayAsync 回傳內容: {dataRequest}");

                    // apiAddTrayAsync 已經回傳正確的格式，直接使用
                    jsonResponse = response;
                }
                catch (Exception parseEx)
                {
                    logger.Error($"Parse response error: {parseEx.Message}, Original response: {dataRequest}");

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "回傳格式解析錯誤: " + parseEx.Message
                    });
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

                logger.Error($"GetLaserEngravedRingCode Error: {e.Message}", e);
            }

            try
            {
                Response.ContentType = "application/json; charset=utf-8";
                Response.Write(jsonResponse.ToString());
            }
            catch (Exception writeEx)
            {
                logger.Error($"Response write error: {writeEx.Message}");
                Response.Write("{\"status\":\"error\",\"msg\":\"回應寫入錯誤\"}");
            }
        }
        #endregion

        #region//AI建議的  晶彩 套環雷刻機台 【API02】紀錄此次雷刻條碼
        [HttpPost]
        [Route("api/MES/apiUpdateLaserEngravedRingCode")]
        public async Task UpdateLaserEngravedRingCode(string Company = "", string SecretKey = "", string UserNo = ""
, string TrayNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateLaserEngravedRingCode");
                #endregion

                #region //Request with Retry - 修正非同步呼叫
                mesBasicInformationDA = new MesBasicInformationDA();

                // 記錄傳入參數用於除錯
                logger.Info($"API02參數 - Company: {Company}, UserNo: {UserNo}, TrayNo: {TrayNo}");

                // 使用重試機制呼叫非同步方法
                dataRequest = await RetryAsync(async () =>
                    await mesBasicInformationDA.UpdateLaserEngravedRingCodeAsync(Company, UserNo, TrayNo),
                    maxRetries: 3);
                #endregion

                #region //Response - 直接解析已格式化的JSON
                try
                {
                    var response = JObject.Parse(dataRequest);
                    logger.Info($"UpdateLaserEngravedRingCodeAsync 回傳內容: {dataRequest}");

                    // UpdateLaserEngravedRingCodeAsync 已經回傳正確的格式，直接使用
                    jsonResponse = response;
                }
                catch (Exception parseEx)
                {
                    logger.Error($"Parse response error: {parseEx.Message}, Original response: {dataRequest}");

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "回傳格式解析錯誤: " + parseEx.Message
                    });
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

                logger.Error($"UpdateLaserEngravedRingCode Error: {e.Message}", e);
            }

            try
            {
                Response.ContentType = "application/json; charset=utf-8";
                Response.Write(jsonResponse.ToString());
            }
            catch (Exception writeEx)
            {
                logger.Error($"Response write error: {writeEx.Message}");
                Response.Write("{\"status\":\"error\",\"msg\":\"回應寫入錯誤\"}");
            }
        }

        #region //重試機制輔助方法 (如果還沒有的話)
        private async Task<T> RetryAsync<T>(Func<Task<T>> operation, int maxRetries = 3, int baseDelayMs = 1000)
        {
            Exception lastException = null;

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    logger.Info($"Attempt {attempt + 1} of {maxRetries}");
                    return await operation();
                }
                catch (SqlException sqlEx) when (IsTransientError(sqlEx))
                {
                    lastException = sqlEx;
                    if (attempt < maxRetries - 1)
                    {
                        int delay = baseDelayMs * (int)Math.Pow(2, attempt); // 指數退避
                        logger.Warn($"Transient error on attempt {attempt + 1}, retrying in {delay}ms: {sqlEx.Message}");
                        await Task.Delay(delay);
                    }
                }
                catch (Exception ex) when (IsNetworkError(ex))
                {
                    lastException = ex;
                    if (attempt < maxRetries - 1)
                    {
                        int delay = baseDelayMs * (attempt + 1);
                        logger.Warn($"Network error on attempt {attempt + 1}, retrying in {delay}ms: {ex.Message}");
                        await Task.Delay(delay);
                    }
                }
            }

            throw lastException ?? new Exception("重試失敗");
        }

        private bool IsTransientError(SqlException ex)
        {
            // SQL Server 暫時性錯誤代碼
            var transientErrors = new[] { 2, 20, 64, 233, 10053, 10054, 10060, 40197, 40501, 40613 };
            return transientErrors.Contains(ex.Number);
        }

        private bool IsNetworkError(Exception ex)
        {
            return ex is TimeoutException ||
                   ex is SocketException ||
                   (ex.InnerException != null && IsNetworkError(ex.InnerException));
        }
        #endregion
        #endregion

        #region//AI建議的  晶彩 套環雷刻機台 【API03】取得雷射刻印人員
        [HttpPost]
        [Route("api/MES/apiGetLaserEngravedUser")]
        public void GetLaserEngravedUser(string Company = "", string SecretKey = "", string LoginNo = "", string PassWord = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetLaserEngravedUser");
                #endregion

                #region //Request
                mesBasicInformationDA = new MesBasicInformationDA();
                dataRequest = mesBasicInformationDA.GetLaserEngravedUser(Company, LoginNo, PassWord);
                #endregion

                #region //Response - 修正回傳格式符合API文件規格
                try
                {
                    var response = JObject.Parse(dataRequest);

                    if (response["status"]?.ToString() == "success")
                    {
                        // 從原始回傳中提取資料
                        var originalData = response["data"];

                        jsonResponse = new JObject
                        {
                            ["status"] = "success",
                            ["msg"] = "OK",
                            ["result"] = originalData // 直接使用原始資料
                        };
                    }
                    else
                    {
                        // 保持原錯誤格式
                        jsonResponse = response;
                    }
                }
                catch (Exception parseEx)
                {
                    logger.Error($"Parse response error: {parseEx.Message}, Original response: {dataRequest}");

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "回傳格式解析錯誤: " + parseEx.Message
                    });
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

                logger.Error($"GetLaserEngravedUser Error: {e.Message}", e);
            }

            try
            {
                Response.ContentType = "application/json; charset=utf-8";
                Response.Write(jsonResponse.ToString());
            }
            catch (Exception writeEx)
            {
                logger.Error($"Response write error: {writeEx.Message}");
                Response.Write("{\"status\":\"error\",\"msg\":\"回應寫入錯誤\"}");
            }
        }
        #endregion
        #endregion
    }
}