using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Text;
using System.Net.Http;
using System.Threading;
using System.Net;
using InterfaceLib;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Tcp;
using MESDA;
using QMSDA;
using System.Diagnostics;
using System.Net.Sockets;
using System.Threading.Tasks;


namespace Business_Manager.Controllers
{
    public class ProductionHistoryController : WebController
    {
        private const string DataPath = @"~/QcDataPath/";
        private ProductionHistoryDA productionHistoryDA = new ProductionHistoryDA();
        private AbnormalqualityDA abnormalqualityDA = new AbnormalqualityDA();

        #region //View
        public ActionResult Index()
        {
            ProductionHistoryLoginCheck();

            return View();
        }
        public ActionResult Login()
        {
            return View();
        }

        public ActionResult LoginSupplier()
        {
            return View();
        }

        public ActionResult MesConsole()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult ProductionHistory()
        {
            ProductionHistoryLoginCheck();

            return View();
        }
        public ActionResult MesScanningInterface()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult MesScanningInterfaceSupplier()
        {

            return View();
        }
        public ActionResult MesIPQC()
        {
            ProductionHistoryLoginCheck();

            return View();
        }
        public ActionResult MesIPQCConsole()
        {
            ProductionHistoryLoginCheck();

            return View();
        }
        public ActionResult MesOQC()
        {
            ProductionHistoryLoginCheck();

            return View();
        }
        public ActionResult MesOQCConsole()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult ToolChoose()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult ToolTrade()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult LetteringChangeFunction()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult AltimeterConsole()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult AltimeterCollection()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult ViewNgBom()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult ProductionHistoryForSupplier()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult indexForSupplier()
        {
            ProductionHistoryLoginCheck();

            return View();

        }

        public ActionResult ViewMachineConsume()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult VehicleManagement()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult ViewMachineSchedule()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult BarcodeBatchIndex()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult BarcodeBatchProcessing()
        {
            ProductionHistoryLoginCheck();

            return View();
        }

        public ActionResult BarcodeMergeProcessing()
        {
            ProductionHistoryLoginCheck();

            return View();

        }

        public ActionResult MaterialFeedReg()
        {
            ProductionHistoryLoginCheck();

            return View();

        }

        public ActionResult MaterialFeedingRecord()
        {
            ProductionHistoryLoginCheck();

            return View();

        }

        #endregion

        #region //Get
        #region //MES過站APP
        #region //查詢製令
        [HttpPost]
        public void GetManufactureOrder(int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetManufactureOrder(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //取得設備IP
        [HttpPost]
        public void GetDeviceIp()
        {
            try
            {
                ProductionHistoryLoginCheck();
                productionHistoryDA = new ProductionHistoryDA();

                string DeviceIdentifierCode = BaseHelper.ClientIP();

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    data = DeviceIdentifierCode
                });
                dataRequest = jsonResponse.ToString();
                jsonResponse = BaseHelper.DAResponse(dataRequest);
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMachineByDevice 取得機台資訊
        [HttpPost]
        public void GetMachineByDevice(string DeviceIdentifierCode = "", int ShopId = -1, string MachineName = "", string MachineNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request 
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMachineByDevice(DeviceIdentifierCode, ShopId, MachineName, MachineNo
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

        #region //GetWorkShopByDevice 取得車間資料-設備用
        [HttpPost]
        public void GetWorkShop(int CompanyId = -1, int ShopId = -1, string ShopNo = "", string ShopName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetWorkShop(CompanyId, ShopId, ShopNo, ShopName, Status
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

        #region //GetMoByProcess 製令對應製程資訊
        [HttpPost]
        public void GetMoByProcess(int MoId = -1, int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request 
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoByProcess(MoId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcessForVehicle 製令對應製程資訊(物料載盤用)
        [HttpPost]
        public void GetMoProcessForVehicle(int MoId = -1, int MachineId = -1, string WoErpFullNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request 
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessForVehicle(MoId, MachineId, WoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcess 取得製程資訊
        [HttpPost]
        public void GetProcess(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request   
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetProcess(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMachine 取得機台資訊
        [HttpPost]
        public void GetMachine(int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request         
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMachine(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUser 取得員工資訊
        [HttpPost]
        public void GetUser(int UserId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetUser(UserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessQty 取得待加工數
        [HttpPost]
        public void GetProcessQty(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request   
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetProcessQty(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMfgOrderProcess 取得製令生產製程設定
        [HttpPost]
        public void GetMfgOrderProcess(int MoId = -1, int MoProcessId = -1, int ModeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request     
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMfgOrderProcess(MoId, MoProcessId, ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetActiveBarcode 取得尚未完工條碼
        [HttpPost]
        public void GetActiveBarcode(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetActiveBarcode(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLetteringInfo 取得刻字相關資訊
        [HttpPost]
        public void GetLetteringInfo(int MoProcessId = -1, string ItemNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetLetteringInfo(MoProcessId, ItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcessId 取得製令歷程資料ID
        [HttpPost]
        public void GetBarcodeProcessId(int BarcodeId = -1, int MoProcessId = -1, string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeProcessId(BarcodeId, MoProcessId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeLettering 取得刻字號
        [HttpPost]
        public void GetBarcodeLettering(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeLettering(BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcessItem 取得製程屬性設定資料
        [HttpPost]
        public void GetMoProcessItem(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessItem(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcessAttribute 取得屬性綁定設定資料
        [HttpPost]
        public void GetBarcodeProcessAttribute(int MoProcessId = -1, string ItemNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeProcessAttribute(MoProcessId, ItemNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetJigBarcode 取得治具綁定資料
        [HttpPost]
        public void GetJigBarcode(int MoProcessId = -1, string JigNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetJigBarcode(MoProcessId, JigNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetJigBarcodeDetail 取得治具綁定條碼資料
        [HttpPost]
        public void GetJigBarcodeDetail(int JigId = -1, string JigNo = "", int MoProcessId = -1, string WorkingFlag = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetJigBarcodeDetail(JigId, JigNo, MoProcessId, WorkingFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcessEvent 取得產品事件資料
        [HttpPost]
        public void GetBarcodeProcessEvent(int MoProcessId = -1, string BarcodeNo = "", int ModeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeProcessEvent(MoProcessId, BarcodeNo, ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoMtlSettingForMTL 取得當前條碼上料相關設定資料
        [HttpPost]
        public void GetMoMtlSettingForMTL(int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoMtlSettingForMTL(MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoMtlSettingForMTL02 取得當前條碼上料相關設定資料
        [HttpPost]
        public void GetMoMtlSettingForMTL02(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoMtlSettingForMTL02(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoMtlSettingForMoId 取得當前製令上料相關設定資料
        [HttpPost]
        public void GetMoMtlSettingForMoId(int MoId = -1, int MoProcessId = -1, string AssemblyBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoMtlSettingForMoId(MoId, MoProcessId, AssemblyBarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcessMaterials 取得當前製令製程需上料之物料
        [HttpPost]
        public void GetMoProcessMaterials(int MoId = -1, int MoProcessId = -1, string AssemblyBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessMaterials(MoId, MoProcessId, AssemblyBarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeComponent 取得被綁定之物料資料
        [HttpPost]
        public void GetBarcodeComponent(int MoId = -1, int MtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeComponent(MoId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeComponentL 取得上料批號模式資料
        [HttpPost]
        public void GetBarcodeComponentL(string AssemblyBarcodeNo  = "", int ComponentMtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeComponentL(AssemblyBarcodeNo, ComponentMtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeComponentByLN 取得上料批號模式資料
        [HttpPost]
        public void GetBarcodeComponentByLN(string AssemblyBarcodeNo = "", int ComponentMtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeComponentByLN(AssemblyBarcodeNo, ComponentMtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoMtlSettingAttribute 取得物料屬性設定資料
        [HttpPost]
        public void GetMoMtlSettingAttribute(int MoId = -1, int MtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoMtlSettingAttribute(MoId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeComponentAttribute 取得物料屬性綁定資料
        [HttpPost]
        public void GetBarcodeComponentAttribute(int MoId = -1, int MtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeComponentAttribute(MoId, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcess 取得製令製程資訊
        [HttpPost]
        public void GetMoProcess(int MoProcessId = -1, int MoId = -1, string WoErpFullNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcess(MoProcessId, MoId, WoErpFullNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLabelPrintMachine 取得標籤機列表
        [HttpPost]
        public void GetLabelPrintMachine()
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetLabelPrintMachine();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetKeyence 取得Keyence設備資料
        [HttpPost]
        public void GetKeyence(int KeyenceId = -1, string Status = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetKeyence(KeyenceId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcess 取得製令條碼歷程
        [HttpPost]
        public void GetBarcodeProcess(int BarcodeProcessId = -1, int MoId = -1, int BarcodeId = -1, string BarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeProcess(BarcodeProcessId, MoId, BarcodeId, BarcodeNo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMoProcessTransaction 取得數量過站歷程
        [HttpPost]
        public void GetMoProcessTransaction(int TransactionId = -1, int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessTransaction(TransactionId, MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessMachine 取得製程參數機台相關資訊
        [HttpPost]
        public void GetProcessMachine(int MoProcessId = -1, int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetProcessMachine(MoProcessId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeInfo 取得指定條碼資訊(批量)
        [HttpPost]
        public void GetBarcodeInfo(string BarcodeList = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeInfo(BarcodeList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcode 取得條碼資料
        [HttpPost]
        public void GetBarcode(string BarcodeNo = "", int MoId = -1, int MoProcessId = -1, int PrintCount = -1, string BarcodeStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcode(BarcodeNo, MoId, MoProcessId, PrintCount, BarcodeStatus
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

        #region //GetMoProcessQty 取得製令製程過站數量資料(數量過站產條碼用)
        [HttpPost]
        public void GetMoProcessQty(int MoProcessId = -1, int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessQty(MoProcessId, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeMoprocess 取得模仁條碼製程資料
        [HttpPost]
        public void GetBarcodeMoprocess(int NextMoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeMoprocess(NextMoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message); 
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPackageBarcodes -- 取得包裝條碼資料
        [HttpPost]
        public void GetPackageBarcodes(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetPackageBarcodes(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPackageBarcodes -- 取得包裝條碼資料
        [HttpPost]
        public void GetPackageProductBarcodes(int PackageBarcodeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetPackageProductBarcodes(PackageBarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCollectionStatus --取得加工事件收集狀態
        [HttpPost]
        public void GetCollectionStatus(int ModeId = -1, int MoProcessId = -1, string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetCollectionStatus(ModeId, MoProcessId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUnfinishManEvent --取得該人員未完成事件
        [HttpPost]
        public void GetUnfinishManEvent()
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetUnfinishManEvent();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeEven --取得該條碼加工事件
        [HttpPost]
        public void GetBarcodeEven(int ModeId = -1, int MoProcessId = -1, string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeEven(ModeId, MoProcessId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUnfinishBarcodeEven --取得該條碼未完成事件
        [HttpPost]
        public void GetUnfinishBarcodeEven(int ModeId = -1, int MoProcessId = -1, string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetUnfinishBarcodeEven(ModeId, MoProcessId, BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUnfinishMachineEvent --取得該條碼未完成事件
        [HttpPost]
        public void GetUnfinishMachineEvent(int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetUnfinishMachineEvent(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeLetteringChangeInfo 取得刻號置換紀錄
        [HttpPost]
        public void GetBarcodeLetteringChangeInfo(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeLetteringChangeInfo(BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeRelevantData NG條碼改刻字，查詢條碼相關資訊
        [HttpPost]
        public void GetBarcodeRelevantData(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeRelevantData(BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeTracking 取得條碼追蹤資料
        [HttpPost]
        public void GetBarcodeTracking(int BarcodeId = -1, string BarcodeNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeTracking(BarcodeId, BarcodeNo
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

        #region //GetAllowQcBarcode 取得可送測條碼資訊
        [HttpPost]
        public void GetAllowQcBarcode(int MoProcessId = 1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetAllowQcBarcode(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMinBarcodes 取得條碼資料
        [HttpPost]
        public void GetMinBarcodes(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMinBarcodes(BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeQty 取得條碼數量
        [HttpPost]
        public void GetBarcodeQty(string BarcodeNo = "", int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeQty(BarcodeNo, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMaterials 取得需上料之物料資料
        [HttpPost]
        public void GetMaterials(int MoProcessId = -1, string VehicleNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMaterials(MoProcessId, VehicleNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetVehicleDetail 取得載盤詳細資料
        [HttpPost]
        public void GetVehicleDetail(string VehicleNo = "", int MoProcessId = -1, string BindType = "", int MtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetVehicleDetail(VehicleNo, MoProcessId, BindType, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCheckProcessJump 確認製程是否有前製程尚未完成
        [HttpPost]
        public void GetCheckProcessJump(int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetCheckProcessJump(MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRoutingItemQcItem 取得途程量測資料
        [HttpPost]
        public void GetRoutingItemQcItem(int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetRoutingItemQcItem(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcReceiptTemplate 取得量測樣板
        [HttpPost]
        public void GetQcReceiptTemplate(int TemplateId = -1, int ModeId = -1, string TemplateName = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetQcReceiptTemplate(TemplateId, ModeId, TemplateName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcReceiptTemplateDetail 取得量測樣板詳細資料
        [HttpPost]
        public void GetQcReceiptTemplateDetail(int DetailId = -1, int TemplateId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetQcReceiptTemplateDetail(DetailId, TemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMachineConsumeLog 取得機台點膠資料
        [HttpPost]
        public void GetMachineConsumeLog(int MoId = -1, int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMachineConsumeLog(MoId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetInProcessBarcode -- 取得在製條碼 
        [HttpPost]
        public void GetInProcessBarcode(string WoErpFullNo = "", int MoProcessId = -1, string StartDate = "", string FinishDate = ""
            , string OrderBy  ="", int PageIndex  =-1, int PageSize = -1)
        {
            try
            {
               // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetInProcessBarcode(WoErpFullNo, MoProcessId, StartDate, FinishDate, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetWorOrderMaterial -- 取得制令所属工单材料
        [HttpPost]
        public void GetWorOrderMaterial(int MoId = -1)
        {
            try
            {
                // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetWorOrderMaterial(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMaterialFeedingRecord -- 取得投料记录
        [HttpPost]
        public void GetMaterialFeedingRecord(int MatFeedRegId = -1, int MachineId = -1, int MoId = -1, string FeedDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMaterialFeedingRecord(MatFeedRegId, MachineId, MoId, FeedDate
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

        #region //GetKeyenceProcessTemp -- 取得條碼暫存資料
        [HttpPost]
        public void GetKeyenceProcessTemp(string TrayNo = "")
        {
            try
            {
                // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetKeyenceProcessTemp(TrayNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //工程檢
        #region //GetFolder 取得資料夾項目
        [HttpPost]
        public void GetFolder(string FilePathSetting = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                string folderRoot = Server.MapPath(DataPath);
                folderRoot = Path.Combine(folderRoot, FilePathSetting);
                //string folderRoot = @"\\192.168.20.199\d8-系統開發室\廖安宏\" + MoId.ToString() + @"\";
                //權限問題 => IIS 識別
                string[] Directories = Directory.GetDirectories(folderRoot);

                List<Folder> folders = new List<Folder>();

                foreach (var d in Directories)
                {
                    var dir = new DirectoryInfo(d);
                    var dirName = dir.Name;
                    var dirPath = dir.ToString();
                    var folderInfo = new Folder
                    {
                        FolderName = dir.Name,
                        FolderPath = dir.ToString()
                    };
                    folders.Add(folderInfo);
                }

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    result = folders
                });
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetFiles 搜尋資料夾內檔案
        [HttpPost]
        public void GetFiles(string FolderPath = "", int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                if (FolderPath == "" && MoId != -1)
                {
                    FolderPath = Server.MapPath(DataPath);
                    FolderPath = Path.Combine(FolderPath, MoId.ToString());
                }
                string ServerFolderPath = Server.MapPath(DataPath);
                FolderPath = Path.Combine(ServerFolderPath, FolderPath);
                string[] Files = Directory.GetFiles(FolderPath);

                List<FileInfoModal> filePaths = new List<FileInfoModal>();

                foreach (var files in Files)
                {
                    var file = new FileInfo(files);

                    var fileInfo = new FileInfoModal
                    {
                        FileName = file.Name,
                        FilePath = file.ToString()
                    };

                    filePaths.Add(fileInfo);
                }

                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    result = filePaths,
                });
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcessIPQC 取得條碼過站歷程資訊
        [HttpPost]
        public void GetBarcodeProcessIPQC(string BarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetBarcodeProcessIPQC(BarcodeNo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetIpqcQcNoticeItemSpec 取得工程檢量測項目及標準值公差
        [HttpPost]
        public void GetIpqcQcNoticeItemSpec(int MoProcessId = -1, string QcNoticeType = "", string Status = "", int QcNoticeItemSpecId = -1, int QcItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetIpqcQcNoticeItemSpec(MoProcessId, QcNoticeType, Status, QcNoticeItemSpecId, QcItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetBarcodeProcessSeq 取得該站可工程檢數量
        [HttpPost]
        public void GetBarcodeProcessSeq(int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetBarcodeProcessSeq(MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTempQcMeasureData 取得工程檢暫存資料
        [HttpPost]
        public void GetTempQcMeasureData(string BarcodeNo = "", int MoProcessId = -1, string Status = "", int MoId = -1, string QcNoticeType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetTempQcMeasureData(BarcodeNo, MoProcessId, Status, MoId, QcNoticeType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcRecordProcess 取得工程檢歷程資料
        [HttpPost]
        public void GetQcRecordProcess(int MoId = -1, int MoProcessId = -1, string BarcodeNo = "", int QcNoticeItemSpecId = -1, int MakeCount = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQcRecordProcess(MoId, MoProcessId, BarcodeNo, QcNoticeItemSpecId, MakeCount
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

        #region //GetQcMeasureDataSeq 取得工程檢項目歷程次數
        [HttpPost]
        public void GetQcMeasureDataSeq(int MoId = -1, int MoProcessId = -1, string BarcodeNo = "", int QcNoticeItemSpecId = -1, int MakeCount = -1, string Cavity = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQcMeasureDataSeq(MoId, MoProcessId, BarcodeNo, QcNoticeItemSpecId, MakeCount, Cavity);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProcessBarcodeInfo 取得加工條碼資訊
        [HttpPost]
        public void GetProcessBarcodeInfo(string BarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetProcessBarcodeInfo(BarcodeNo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetFilePathSetting 取得量測上傳路徑
        [HttpPost]
        public void GetFilePathSetting(int QcMachineModeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetFilePathSetting(QcMachineModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcFile 取得量測上傳路徑
        [HttpPost]
        public void GetQcFile(int QcFileId = -1, int QcNoticeId = -1, int QcNoticeItemSpecId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQcFile(QcFileId, QcNoticeId, QcNoticeItemSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQmmDetail 取得量測機台詳細資料
        [HttpPost]
        public void GetQmmDetail(int QmmDetailId = -1, int MachineId = -1, string MachineNumber = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQmmDetail(QmmDetailId, MachineId, MachineNumber);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //出貨檢
        #region //GetOQcBarcode 取得出貨檢條碼資訊
        [HttpPost]
        public void GetOQcBarcode(string BarcodeNo = "", int MoId = -1, string QcType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetOQcBarcode(BarcodeNo, MoId, QcType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetOQcMeasureDataSeq 取得工程檢項目歷程次數
        [HttpPost]
        public void GetOQcMeasureDataSeq(int MoId = -1, string BarcodeNo = "", int QcNoticeItemSpecId = -1, int MakeCount = -1, string Cavity = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetOQcMeasureDataSeq(MoId, BarcodeNo, QcNoticeItemSpecId, MakeCount, Cavity);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetOqcQcNoticeItemSpec 取得出貨檢量測項目
        [HttpPost]
        public void GetOqcQcNoticeItemSpec(int MoId = -1, string QcNoticeType = "", string Status = "", int QcNoticeItemSpecId = -1, int QcItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetOqcQcNoticeItemSpec(MoId, QcNoticeType, Status, QcNoticeItemSpecId, QcItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //全吋檢
        #region //GetTqcQcNoticeItemSpec 取得出貨檢量測項目
        [HttpPost]
        public void GetTqcQcNoticeItemSpec(int MoId = -1, string QcNoticeType = "", string Status = "", int QcNoticeItemSpecId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetTqcQcNoticeItemSpec(MoId, QcNoticeType, Status, QcNoticeItemSpecId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //高度計
        #region //GetBarcodeProcessBarcode 取得該站可工程檢條碼
        [HttpPost]
        public void GetBarcodeProcessBarcode(int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetBarcodeProcessBarcode(MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region GetQcNoticeQcItem//取得編程資料量測項目
        [HttpPost]
        public void GetQcNoticeQcItem(int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQcNoticeQcItem(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetQcAltimeter --取得量測高度計機台
        [HttpPost]
        public void GetQcAltimeter(string DeviceIdentifierCode = "", int ShopId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQcAltimeter(DeviceIdentifierCode, ShopId
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

        #region GetQcNoticeQcItem//取得編程資料量測項目
        [HttpPost]
        public void GetAltimeterData(int MoId = -1, int MoProcessId = -1, string MeasurementNo = "") 
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetAltimeterData(MoId, MoProcessId, MeasurementNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region GetQcNoticeQcItemValue////取得量測項目設定值
        [HttpPost]
        public void GetQcNoticeQcItemValue(int QcNoticeId = -1, int QcItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.GetQcNoticeQcItemValue(QcNoticeId, QcItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //物料載盤
        #region //GetVehicle 取得載盤資料
        [HttpPost]
        public void GetVehicle(string VehicleNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetVehicle(VehicleNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //MES過站資料收集之用途
        #region //AddBarcodeProcessAttribute 新增製程屬性資料
        [HttpPost]
        public void AddBarcodeProcessAttribute(string BarcodeProcessAttributeData = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeProcessAttribute(BarcodeProcessAttributeData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddJigBarcode 新增製令治具
        [HttpPost]
        public void AddJigBarcode(int JigId = -1, string JigNo = "", string BarcodeNo = "", int MoProcessId = -1, int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddJigBarcode(JigId, JigNo, BarcodeNo, MoProcessId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddLotNgBarcodeProcess 新增批量NG條碼
        [HttpPost]
        public void AddLotNgBarcodeProcess(string BarcodeNo = "", string NgBarcodeNo = "", int NgBarcodeQty = -1, string CauseId = "", int MoId = -1, int MachineId = -1, string CauseNo = "", int MoProcessId = -1, string MinBarcodes = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddLotNgBarcodeProcess(BarcodeNo, NgBarcodeNo, NgBarcodeQty, CauseId, MoId, MachineId, CauseNo, MoProcessId, MinBarcodes);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBarcodeComponent 新增物料元件資料
        [HttpPost]
        public void AddBarcodeComponent(string AssemblyBarcodeNo = "", string ComponentBarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeComponent(AssemblyBarcodeNo, ComponentBarcodeNo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBarcodeComponentL 新增物料元件資料(批號)
        [HttpPost]
        public void AddBarcodeComponentL(string AssemblyBarcodeNo = "", string LotNumber = "", int MoProcessId = -1, int ComponentMtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeComponentL(AssemblyBarcodeNo, LotNumber, MoProcessId, ComponentMtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddEmptyBarcode 新增批量空良品條碼
        [HttpPost]
        public void AddEmptyBarcode(string FromBarcodeNo = "", int MoId = -1, string NewTrayBarcode = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddEmptyBarcode(FromBarcodeNo, MoId, NewTrayBarcode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBindBarcode 新增條碼(數量過站產生條碼用)
        [HttpPost]
        public void AddBindBarcode(int MoProcessId = -1, string BarcodePrefix = "", string BarcodePostfix = "", int SequenceLen = -1, int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBindBarcode(MoProcessId, BarcodePrefix, BarcodePostfix, SequenceLen, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBarcodeComponentAttribute 新增上料元件屬性資料
        [HttpPost]
        public void AddBarcodeComponentAttribute(string BarcodeComponentAttributeData = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeComponentAttribute(BarcodeComponentAttributeData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddLotBarcodeQtyMerge 新增批量條碼拆併盤資料
        [HttpPost]
        public void AddLotBarcodeQtyMerge(int MoId = -1, string FromBarcodeNo = "", string ToBarcodeNo = "", int TransactionQty = -1, string MinBarcodes = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddLotBarcodeQtyMerge(MoId, FromBarcodeNo, ToBarcodeNo, TransactionQty, MinBarcodes);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddQcRecord 新增量測紀錄單頭
        [HttpPost]
        public void AddQcRecord(int QcNoticeId = -1, int QcTypeId = -1, int MoId = -1, int MoProcessId = -1, int QmmDetailId = -1, int QcItemId = -1, string Remark = "", int CurrentFileId = -1, string QcDepartment = "", string MachineNumber = ""
            , string LotNumber = "", string QcRecordFile = "", string InputType = "", string ServerPath = "", string ServerPath2 = "", string CheckQcMeasureData = "", string SupportAqFlag = "", string ResolveFile = ""
            , string BarcodeList = "", string ResolveFileJson = "", string QcRecordFileByNas = "", string QcMeasureDataJson = "", string SourcePage = "", string QcItemJson = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //建立量測單據
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddQcRecord(QcNoticeId, QcTypeId, MoId, MoProcessId, QmmDetailId, QcItemId, Remark, CurrentFileId, QcDepartment, MachineNumber
                    , LotNumber, QcRecordFile, InputType, "", "", CheckQcMeasureData, SupportAqFlag, ResolveFile
                    , BarcodeList, ResolveFileJson, QcRecordFileByNas, QcMeasureDataJson, SourcePage, QcItemJson);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPackageBarcode --新增包裝條碼
        [HttpPost]
        public void AddPackageBarcode(int MoProcessId = -1, int MoId = -1, string CreatDate = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddPackageBarcode(MoProcessId, MoId, CreatDate); 
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPackageProductBarcode --新增包裝產品條碼
        [HttpPost]
        public void AddPackageProductBarcode(int PackageBarcodeId = -1, string BarcodeList = "", int MoId = -1, int CurrenMoProcess = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddPackageProductBarcode(PackageBarcodeId, BarcodeList, MoId, CurrenMoProcess);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddStartCoutTime --新增事件開始資料(機台/加工)
        [HttpPost]
        public void AddStartCoutTime(int EventItemId = -1, int BarcodeId = -1, int MoProcessId = -1, string StartTime = "", string CountType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddStartCoutTime(EventItemId, BarcodeId, MoProcessId, StartTime, CountType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddManStartCoutTime --新增事件開始資料(人員)
        [HttpPost]
        public void AddManStartCoutTime(int EventItemId = -1, string StartTime = "", string CountType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddManStartCoutTime(EventItemId, StartTime, CountType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMaterialFeedReg 新增投料登记 
        [HttpPost]
        public void AddMaterialFeedReg(int MachineId = -1, int MoId = -1, int MaterialId = -1, double MatInRegQty = -1, int UomId = -1, string Remarks = "")
        {
            try
            {
               // WebApiLoginCheck("ItemCategoryDept", "add");

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddMaterialFeedReg(MachineId, MoId, MaterialId, MatInRegQty , UomId, Remarks);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddKeyenceProcessTemp --新增KeyenceTemp
        [HttpPost]
        public void AddKeyenceProcessTemp(string BarcodeList = "", int MoProcessId = -1, int MachineId = -1, int UserId = -1, int KeyenceId = -1
            , string QtyBarcodeStatus = "", string QtyLotNgCauseNo = "", int QtyNextMoProcessId = -1, string ItemNo = "", string ItemValue = "", string ChkUnique = "", string DateCode = "")
        {
            try
            {
                
                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddKeyenceProcessTemp(BarcodeList, MoProcessId, MachineId, UserId, KeyenceId
                    , QtyBarcodeStatus, QtyLotNgCauseNo, QtyNextMoProcessId, ItemNo, ItemValue, ChkUnique, DateCode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //工程檢
        #region //AddTempQcMeasureData 新增工程檢量測備份資料
        [HttpPost]
        public void AddTempQcMeasureData(string TempQcItemData = "", int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.AddTempQcMeasureData(TempQcItemData, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //出貨檢/試產檢
        #region //AddTempOQcMeasureData 新增出貨檢量測備份資料
        [HttpPost]
        public void AddTempOQcMeasureData(string TempOQcItemData = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.AddTempOQcMeasureData(TempOQcItemData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBarcodeOQcRecord 新增出貨檢紀錄 + 上傳量測數據 + 自動建立品異單
        [HttpPost]
        public void AddBarcodeOQcRecord(string BarcodeQcRecordData = "", int MoId = -1, string TempOQcItemData = "", string QcType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.AddBarcodeOQcRecord(BarcodeQcRecordData, MoId, TempOQcItemData, QcType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //全吋檢
        #region //AddBarcodeTQcRecord 新增全吋檢紀錄 + 上傳量測數據
        [HttpPost]
        public void AddBarcodeTQcRecord(string BarcodeQcRecordData = "", int MoId = -1, string TempOQcItemData = "", string QcType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.AddBarcodeTQcRecord(BarcodeQcRecordData, MoId, TempOQcItemData, QcType);
                #endregion

                #region //更改暫存資料狀態
                var dataRequestJson = JObject.Parse(dataRequest);
                if (dataRequest.IndexOf("errorForDA") != -1)
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }

                foreach (var item in dataRequestJson)
                {
                    if (item.Key == "status")
                    {
                        if (item.Value.ToString() != "success") throw new SystemException("error");
                    }
                }

                dataRequest = productionHistoryDA.UpdateTempOQcMeasureData(MoId, "", "S", -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //品異
        #region //AddAbnormalqualityPadProject 新增品異單 - 平板版本(到對策確認前)
        [HttpPost]
        public void AddAbnormalqualityPadProject(string AbnormalqualityData = "", string AbnormalProjectList = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.AddAbnormalqualityPadProject(AbnormalqualityData, AbnormalProjectList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //高度計

        #region //AddAltimeterData 新增高度計數據
        [HttpPost]
        public void AddAltimeterData(int MoId = -1, int MoProcessId = -1, int QcNoticeId = -1, int MachineId = -1, int BarcodeId = -1, int QcItemId = -1, string QcStatus = "", float MeasureValue = -1, float UpperValue = -1, float LowerValue = -1, float DesignValue = -1, int ModeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.AddAltimeterData(MoId, MoProcessId, QcNoticeId, MachineId, BarcodeId, QcItemId, QcStatus, MeasureValue, UpperValue, LowerValue, DesignValue, ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //物料載盤
        #region //AddVehicleDetail 新增物料載盤詳細資料
        [HttpPost]
        public void AddVehicleDetail(string VehicleNo = "", int MoProcessId = -1, string BarcodeNo = "", string LotNumberNo = "", string BindType = "", int MtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddVehicleDetail(VehicleNo, MoProcessId, BarcodeNo, LotNumberNo, BindType, MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Update
        #region //MES過站資料收集之用途
        #region //TxBarcodeProcess -- BARCODE DATA COLLECT(目前過站主程式)
        [HttpPost]
        public void TxBarcodeProcess(string BarcodeNo = "", int MoProcessId = -1, int MachineId = -1, int UserId = -1,
            string QtyBarcodeStatus = "", string QtyLotNgCauseNo = "", int QtyNextMoProcessId = -1, string Company = "", string UserNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.TxBarcodeProcess(BarcodeNo, MoProcessId, MachineId, UserId
                    , QtyBarcodeStatus, QtyLotNgCauseNo, QtyNextMoProcessId, Company, UserNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //TxBarcodeAttributeLog -- 現場NG模仁改刻字 (模仁改刻字置換介面)
        [HttpPost]
        public void TxBarcodeAttributeLog(string BarcodeNo = "", int UserId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.TxBarcodeAttributeLog(BarcodeNo, UserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //TxKeyenceProcess -- Keyence模式過站入口
        [HttpPost]
        public void TxKeyenceProcess(string BarcodeList = "", int MoProcessId = -1, int MachineId = -1, int UserId = -1, int KeyenceId = -1
            , string QtyBarcodeStatus = "", string QtyLotNgCauseNo = "", int QtyNextMoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.TxKeyenceProcess(BarcodeList, MoProcessId, MachineId, UserId, KeyenceId, QtyBarcodeStatus, QtyLotNgCauseNo, QtyNextMoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //TxOspBarcodeProcess -- 託外過站主程式
        [HttpPost]
        public void TxOspBarcodeProcess(int OsrId = -1, string SourceMode = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.TxOspBarcodeProcess(OsrId, SourceMode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //TxKeyenceProcess2 -- Keyence模式過站入口(單次)
        [HttpPost]
        public void TxKeyenceProcess2(string Barcode = "", int MoProcessId = -1, int MachineId = -1, int UserId = -1, int KeyenceId = -1
            , string QtyBarcodeStatus = "", string QtyLotNgCauseNo = "", int QtyNextMoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.TxKeyenceProcess(Barcode, MoProcessId, MachineId, UserId, KeyenceId, QtyBarcodeStatus, QtyLotNgCauseNo, QtyNextMoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //TxKeyenceProcessFromTemp -- Keyence模式過站入口(暫存新增)
        [HttpPost]
        [Route("api/MES/TxKeyenceProcessFromTemp")]
        public async Task TxKeyenceProcessFromTemp(int chunkSize = 100)
        {
            try
            {
                //ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = await productionHistoryDA.TxKeyenceProcessFromTemp(chunkSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //UpdateJigBarcodeWorkingFlag 更新治具條碼完工狀態
        [HttpPost]
        public void UpdateJigBarcodeWorkingFlag(string JigBarcodeIdData = "", string JigNo = "", int MoProcessId = -1, int MachineId = -1, int UserId = -1, string Company = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateJigBarcodeWorkingFlag(JigBarcodeIdData);

                if (dataRequest.IndexOf("success") != -1)
                {
                    string BarcodeStatus = "";
                    string LotNgCauseNo = "";
                    int NextMoProcessId = -1;
                    dataRequest = productionHistoryDA.TxBarcodeProcess(JigNo, MoProcessId, MachineId, UserId, BarcodeStatus, LotNgCauseNo, NextMoProcessId, Company, "");

                    //if (dataRequest.IndexOf("success") != -1)
                    //{
                    //    #region //查詢條碼是否完工，若完工則自動解除模仁治具綁定
                    //    dataRequest = productionHistoryDA.UpdateUnbindJigBarcode(JigBarcodeIdData, JigNo);
                    //    #endregion
                    //}
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

        #region //UpdateLotNgBarcodeProcessToPass 更新批量條碼數量為P
        [HttpPost]
        public void UpdateLotNgBarcodeProcessToPass(int BarcodeProcessId = -1, int ChangeBarcodeQty = -1, string PreBarcodeNo = "", string MinBarcodes = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateLotNgBarcodeProcessToPass(BarcodeProcessId, ChangeBarcodeQty, PreBarcodeNo, MinBarcodes);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateQunatityJudge 更新數量過站資料
        [HttpPost]
        public void UpdateQunatityJudge(int TransactionId = -1, string ToJudgeType = "", int ChangeQty = -1, int NextMoProcessId = -1, string CauseNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateQunatityJudge(TransactionId, ToJudgeType, ChangeQty, NextMoProcessId, CauseNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateBarcodeQty 更新條碼數量(數量產條碼功能用)
        [HttpPost]
        public void UpdateBarcodeQty(int BarcodeId = -1, int BarcodeQty = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateBarcodeQty(BarcodeId, BarcodeQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateNextMoProcess 更新條碼下一站製程
        [HttpPost]
        public void UpdateNextMoProcess(string BarcodeNo = "", int MoProcessId = -1, int NewMoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateNextMoProcess(BarcodeNo, MoProcessId, NewMoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateNextMoProcessForLotMode 更新條碼下一站製程(批量修改)
        [HttpPost]
        public void UpdateNextMoProcessForLotMode(string BarcodeNoListString = "", int MoProcessId = -1, int NewMoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateNextMoProcessForLotMode(BarcodeNoListString, MoProcessId, NewMoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdStatus 更新條碼狀態
        [HttpPost]
        public void UpdateProdStatus(int BarcodeProcessId = -1, string ProdStatus = "", string CauseNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateProdStatus(BarcodeProcessId, ProdStatus, CauseNo);
                #endregion

                var dataRequestJson = JObject.Parse(dataRequest);

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProdStatusForLotMode 更新條碼狀態(批量修改版本)
        [HttpPost]
        public void UpdateProdStatusForLotMode(string BarcodeProcessIdListString = "", string ProdStatus = "", string CauseNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateProdStatusForLotMode(BarcodeProcessIdListString, ProdStatus, CauseNo);
                #endregion

                var dataRequestJson = JObject.Parse(dataRequest);

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateEndCoutTime --更新該事件完成時間
        [HttpPost]
        public void UpdateEndCoutTime(int EventId = -1, string Finish = "", string CountType = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateEndCoutTime(EventId, Finish, CountType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //工程檢
        #region //UpdateTempQcMeasureData 更新工程檢量測數據暫存資料
        [HttpPost]
        public void UpdateTempQcMeasureData(int BarcodeId = -1, int MoProcessId = -1, string Status = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.UpdateTempQcMeasureData(BarcodeId, MoProcessId, Status);

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //出貨檢
        #region //UpdateTempOQcMeasureData 更新出貨檢量測數據暫存資料
        [HttpPost]
        public void UpdateTempOQcMeasureData(int MoId = -1, string BarcodeNo = "", string Status = "", int BarcodeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                dataRequest = productionHistoryDA.UpdateTempOQcMeasureData(MoId, BarcodeNo, Status, BarcodeId);

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //物料載盤
        #region //UpdateImportVehicle 物料載盤 MAPPING TrayBarcode
        [HttpPost]
        public void UpdateImportVehicle(string VehicleNo = "", int MoProcessId = -1, string AssemblyBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateImportVehicle(VehicleNo, MoProcessId, AssemblyBarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //條碼入庫後拆併及包裝作業
        #region //BarcodeMergerTransaction 條碼拆併盤
        [HttpPost]
        public void BarcodeMergerTransaction(string FromBarcodeNo = "", string ToBarcodeNo = "", int TransactionQty = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.BarcodeMergerTransaction(FromBarcodeNo, ToBarcodeNo, TransactionQty);
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

        #region //GetBarcodeForBatchProcessing 取得條碼資料(拆盤包裝作業用)
        [HttpPost]
        public void GetBarcodeForBatchProcessing(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeForBatchProcessing(BarcodeNo);
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

        #region //RegisterBarcodeToPackage 註冊產品條碼至包裝中
        [HttpPost]
        public void RegisterBarcodeToPackage(string BarcodeNo = "", string PrevBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.RegisterBarcodeToPackage(BarcodeNo, PrevBarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddEmptyPackage 新增空包裝條碼
        [HttpPost]
        public void AddEmptyPackage()
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddEmptyPackage();
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

        #region //GetPackageBarcode 取得包裝內條碼資料
        [HttpPost]
        public void GetPackageBarcode(string PrevBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetPackageBarcode(PrevBarcodeNo);
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

        #region //UnRegisterBarcodeToPackage 將條碼從包裝中移除註冊
        [HttpPost]
        public void UnRegisterBarcodeToPackage(string BarcodeNo = "", string PrevBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UnRegisterBarcodeToPackage(BarcodeNo, PrevBarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCurrentPackageBarcode 取得目前產品條碼資料
        [HttpPost]
        public void GetCurrentPackageBarcode(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetCurrentPackageBarcode(BarcodeNo);
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
        #endregion

        #region //UpdateBarcodesToScrap -- 批次更改條碼為報廢
        [HttpPost]
        public void UpdateBarcodesToScrap(string BarcodeInfo = "", string ScrapReason = "")
        {
            try
            {
               // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateBarcodesToScrap(BarcodeInfo, ScrapReason);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateLotBarcodeQty -- 修改條碼數量
        [HttpPost]
        public void UpdateLotBarcodeQty(int PrintId = -1, int SetQty = -1)
        {
            try
            {
                // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateLotBarcodeQty(PrintId, SetQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMaterialFeedingRecord -- 修改投料记录
        [HttpPost]
        public void UpdateMaterialFeedingRecord(int MatFeedRegId = -1,  int MaterialId = -1, double MatInRegQty = -1,int UomId = -1,string Remarks = "")
        {
            try
            {
                // ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateMaterialFeedingRecord(MatFeedRegId, MaterialId,  MatInRegQty, UomId, Remarks);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateStatusMaterialFeedingRecord -- 修改投料记录状态为作废
        [HttpPost]
        public void UpdateStatusMaterialFeedingRecord(int MatFeedRegId = -1, string Status = "")
        {
            try
            {
                // ProductionHistoryLoginCheck();
                 
                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateStatusMaterialFeedingRecord(MatFeedRegId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteJigBarcodeDetail -- 刪除治具綁定條碼資料
        [HttpPost]
        public void DeleteJigBarcodeDetail(int JigBarcodeId = -1, string BarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeleteJigBarcodeDetail(JigBarcodeId, BarcodeNo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBarcodeComponent -- 刪除上料綁定資料
        [HttpPost]
        public void DeleteBarcodeComponent(int BarcodeComponentId = -1, string ComponentBarcodeNo = "", int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeleteBarcodeComponent(BarcodeComponentId, ComponentBarcodeNo, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBarcodeComponentL -- 刪除上料綁定資料(批號模式)
        [HttpPost]
        public void DeleteBarcodeComponentL(int BarcodeComponentId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeleteBarcodeComponentL(BarcodeComponentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteAllBarcodeComponent -- 刪除上料綁定資料(批量解綁)
        [HttpPost]
        public void DeleteAllBarcodeComponent(int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeleteAllBarcodeComponent(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteBarcode -- 刪除綁定條碼資料(數量產條碼功能用)
        [HttpPost]
        public void DeleteBarcode(int BarcodeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeleteBarcode(BarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //刪除包裝產品條碼
        [HttpPost]
        public void DeletePackageProductBarcode(int PackageBarcodeId = -1, string BarcodeNo = "", int MoId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeletePackageProductBarcode(PackageBarcodeId, BarcodeNo, MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //刪除包裝條碼
        [HttpPost]
        public void DeletePackageBarcode(int PackageBarcodeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeletePackageBarcode(PackageBarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //物料載盤
        #region //DeleteVehicleDetail -- 解綁物料載盤
        [HttpPost]
        public void DeleteVehicleDetail(int VdId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.DeleteVehicleDetail(VdId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //API
        #region //UploadQcFile -- 上傳量測檔案(工程檢、出貨檢) -- Ann 2022-10-18
        [HttpPost]
        public void UploadQcFile(string PostData = "")
        {
            try
            {
                
                string QCData = "";
                using (HttpClient client = new HttpClient())
                {
                    string apiIP = "http://192.168.20.46:1415/";
                    string apiUrl = "http://192.168.20.46:1415/PmdWebSystem/qa_management_system";
                    client.Timeout = TimeSpan.FromMinutes(60);
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        var postDataJson = JObject.Parse(PostData);

                        MultipartFormDataContent multipart = new MultipartFormDataContent();

                        foreach (var item in postDataJson)
                        {
                            string KValue = item.Value.ToString().Replace(System.Environment.NewLine, string.Empty);
                            KValue = KValue.Trim();
                            KValue = KValue.Replace(" ", "");
                            multipart.Add(new StringContent(KValue), item.Key);
                        }

                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                var data = httpResponseMessage.Content.ReadAsStringAsync();
                                var dataJson = JObject.Parse(data.Result.ToString());
                                string LOG_ID = "";
                                foreach (var item in dataJson)
                                {
                                    if (item.Key.ToString() == "LOG_ID")
                                    {
                                        LOG_ID = item.Value.ToString();
                                        break;
                                    }
                                }

                                Thread.Sleep(2000);

                                string result = "";
                                while (true)
                                {
                                    string Geturl = apiIP + "GetStatusY/" + LOG_ID;//自己指定URL
                                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Geturl);
                                    request.Method = "GET";
                                    request.ContentType = "text/json;charset=UTF-8";
                                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                                    Stream myResponseStream = response.GetResponseStream();
                                    StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
                                    string retString = myStreamReader.ReadToEnd();
                                    myStreamReader.Close();
                                    myResponseStream.Close();

                                    if (retString != "{}\n")
                                    {
                                        result = retString;
                                        break;
                                    }
                                    Thread.Sleep(1000);
                                }

                                var resultJson = JObject.Parse(result);

                                int MoId = -1;
                                int MoProcessId = -1;
                                string ClientIP = "";
                                Boolean errorMessageBool = false;
                                int QcNoticeItemSpecId = -1;
                                string QcNoticeType = "";
                                foreach (var item in resultJson)
                                {
                                    if (item.Key == "moId")
                                    {
                                        MoId = Convert.ToInt32(item.Value);
                                    }
                                    else if (item.Key == "moProcessId")
                                    {
                                        MoProcessId = Convert.ToInt32(item.Value);
                                    }
                                    else if (item.Key == "pc_ip")
                                    {
                                        ClientIP = item.Value.ToString();
                                    }
                                    else if (item.Key == "measureType")
                                    {
                                        QcNoticeType = item.Value.ToString();
                                    }
                                    else if (item.Key == "error_message")
                                    {
                                        if (item.Value.ToString() != "") throw new SystemException(item.Value.ToString());
                                        else errorMessageBool = true;
                                    }

                                    if (MoId > 0 && MoProcessId > 0 && ClientIP != ""  && QcNoticeType != "" && errorMessageBool == true) break;
                                }

                                foreach (var item in resultJson)
                                {
                                    if (item.Key == "data")
                                    {
                                        if (item.Value.ToString() == "[]") throw new SystemException("檔案解析過程有誤!");
                                        //解析回傳量測值並修改相關設定
                                        var QcDataJson = JToken.Parse(item.Value.ToString());
                                        int count = 0;
                                        string barcodeNo = "";
                                        foreach (var item2 in QcDataJson)
                                        {
                                            #region //取得BarcodeInfo
                                            string ItemValue = QcDataJson[count]["Lettering"].ToString();
                                            if (ItemValue != "")
                                            {
                                                var barcodeInfoResult = productionHistoryDA.GetBarcodeForLettering(ItemValue);
                                                var barcodeInfoResultJson = JObject.Parse(barcodeInfoResult);

                                                foreach (var item3 in barcodeInfoResultJson)
                                                {
                                                    if (item3.Key == "data")
                                                    {
                                                        var keyValue = item3.Value;
                                                        if (keyValue.Count() <= 0) throw new SystemException("條碼資訊錯誤!");
                                                        barcodeNo = keyValue[0]["BarcodeNo"].ToString();
                                                        QcDataJson[count]["BarcodeNo"] = barcodeNo;
                                                    }
                                                }
                                            }
                                            #endregion

                                            #region //取得機台ID
                                            string QcMachineNumber = QcDataJson[count]["QcMachineNumber"].ToString();
                                            var QcMachineNoResult = productionHistoryDA.GetQmmDetail(-1, -1, QcMachineNumber);
                                            var QcMachineNoResultJson = JObject.Parse(QcMachineNoResult);

                                            foreach (var item3 in QcMachineNoResultJson)
                                            {
                                                if (item3.Key == "data")
                                                {
                                                    var keyValue = item3.Value;
                                                    if (keyValue.Count() <= 0) throw new SystemException("查無此機台【" + QcMachineNumber + "】資訊!");
                                                    int QcMachineModeId = Convert.ToInt32(keyValue[0]["QcMachineModeId"]);
                                                    int QmmDetailId = Convert.ToInt32(keyValue[0]["QmmDetailId"]);
                                                    string MachineDesc = keyValue[0]["MachineDesc"].ToString();
                                                    QcDataJson[count]["QcMachineModeId"] = QcMachineModeId;
                                                    QcDataJson[count]["QmmDetailId"] = QmmDetailId;
                                                    QcDataJson[count]["MachineDesc"] = MachineDesc;
                                                }
                                            }
                                            #endregion

                                            #region //取得QcNoticeItemSpecId
                                            int qcItemId = Convert.ToInt32(QcDataJson[count]["QcItemId"]);
                                            var QcNoticeItemSpecIdResult = "";

                                            if (QcNoticeType == "IPQC")
                                            {
                                                QcNoticeItemSpecIdResult = productionHistoryDA.GetIpqcQcNoticeItemSpec(MoProcessId, QcNoticeType, "N", -1, qcItemId);
                                            }
                                            else if (QcNoticeType == "OQC" || QcNoticeType == "TVPQC")
                                            {
                                                QcNoticeItemSpecIdResult = productionHistoryDA.GetOqcQcNoticeItemSpec(MoId, QcNoticeType, "N", -1, qcItemId);
                                            }
                                            else
                                            {
                                                throw new SystemException("此量測類型【" + QcNoticeType + "】尚未開發!");
                                            }
                                            var QcNoticeItemSpecIdResultJson = JObject.Parse(QcNoticeItemSpecIdResult);

                                            foreach (var item3 in QcNoticeItemSpecIdResultJson)
                                            {
                                                if (item3.Key == "data")
                                                {
                                                    var keyValue = item3.Value;
                                                    if (keyValue.Count() <= 0) throw new SystemException("量測項目資料錯誤!");
                                                    QcNoticeItemSpecId = Convert.ToInt32(keyValue[0]["QcNoticeItemSpecId"]);
                                                    int DesignValue = Convert.ToInt32(keyValue[0]["DesignValue"]);
                                                    int UpperTolerance = Convert.ToInt32(keyValue[0]["UpperTolerance"]);
                                                    int LowerTolerance = Convert.ToInt32(keyValue[0]["LowerTolerance"]);
                                                    string QcItemType = keyValue[0]["QcItemType"].ToString();
                                                    QcDataJson[count]["QcNoticeItemSpecId"] = QcNoticeItemSpecId;
                                                    QcDataJson[count]["DesignValue"] = DesignValue;
                                                    QcDataJson[count]["UpperTolerance"] = UpperTolerance;
                                                    QcDataJson[count]["LowerTolerance"] = LowerTolerance;
                                                    QcDataJson[count]["QcItemType"] = QcItemType;
                                                }
                                            }
                                            #endregion
                                            count++;
                                        }
                                        QCData = QcDataJson.ToString();
                                        break;
                                    }
                                }

                                foreach (var item in resultJson)
                                {
                                    if (item.Key == "filePaths")
                                    {
                                        var QCDataJson = JToken.Parse(QCData);

                                        int fileCount = 0;
                                        foreach (var item3 in QCDataJson)
                                        {
                                            QCDataJson[fileCount]["FilePath"] = item.Value.ToString();
                                            fileCount++;
                                        }

                                        QCData = QCDataJson.ToString();

                                        var filePathsJson = JToken.Parse(item.Value.ToString());

                                        string count = "0";
                                        foreach (var item2 in filePathsJson)
                                        {
                                            string filePath = item2[count].ToString();
                                            byte[] FileContent = System.IO.File.ReadAllBytes(filePath);
                                            string FileName = Path.GetFileNameWithoutExtension(filePath);
                                            string FileExtension = Path.GetExtension(filePath);
                                            FileInfo fileInfo = new FileInfo(filePath);
                                            int FileSize = Convert.ToInt32(fileInfo.Length);
                                            string Source = "";
                                            if (QcNoticeType == "IPQC") Source = "/ProductionHistory/ViewUploadIPQCFile";
                                            else if (QcNoticeType == "OQC" || QcNoticeType == "TVPQC") Source = "/ProductionHistory/ViewUploadOQCFile";
                                            else throw new SystemException("此量測類型【" + QcNoticeType + "】尚未開發!");
                                            #region //寫入資料庫
                                            var qcFileRequest =  productionHistoryDA.AddQcFile(MoId, MoProcessId, FileContent, FileName, FileExtension, FileSize, ClientIP, Source, QcNoticeItemSpecId);
                                            var qcFileRequestJson = JObject.Parse(qcFileRequest);
                                            if (qcFileRequest.IndexOf("success") == -1)
                                            {
                                                foreach (var item3 in qcFileRequestJson)
                                                {
                                                    if (item3.Key == "msg")
                                                    {
                                                        throw new SystemException(item3.Value.ToString());
                                                    }
                                                }
                                            }
                                            #endregion
                                            count = (Convert.ToInt32(count) + 1).ToString();
                                        }
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                //throw new SystemException(httpResponseMessage.RequestMessage.ToString());
                                throw new SystemException("伺服器連線異常，請聯絡資訊人員!");
                            }
                        }
                    }
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = QCData
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

        #region//UploadNcCode --上拋程式
        [HttpPost]
        public void UploadNcCode(int CompanyId=-1, string UserNo="", int MoProcessId=-1, int MachineId=-1)
        {
            string result = "", error = "";
            try
            {
                #region --執行get File並下載到指定路徑

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetCncProgramWorkFiles(CompanyId, MoProcessId, MachineId);
                jsonResponse = JObject.Parse(dataRequest);
                foreach (var item in jsonResponse["data"])
                {
                    string FileName = item["FileName"].ToString();
                    string FileExtension = item["FileExtension"].ToString();
                    byte[] FileContent = (byte[])item["FileContent"];
                    string NcName = item["NcName"].ToString();
                    if (NcName.Length == 5)
                    {
                        string DocPath = Path.Combine(Server.MapPath(@"~\Uploads\CncProgramWorkFiles\"), FileName + FileExtension);
                        System.IO.File.WriteAllBytes(DocPath, FileContent);                      
                        result = UPLOAD_NC_CODE(MachineId, DocPath + "#" + NcName);
                    }
                    else
                    {
                        logger.Error("操作者:" + UserNo + "，異常回報:不符合編程命名規範");
                    }
                }

                #region//刪除            
                DirectoryInfo di = new DirectoryInfo(Server.MapPath(@"~\Uploads\CncProgramWorkFiles\"));
                FileInfo[] files = di.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
                #endregion

                #endregion
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
            Response.Write(result);
        }
        #endregion

        #region //UploadNcCodetoMachine --上拋機台
        public void UploadNcCodetoMachine(string FilePath, int MachineId, string UserNo)
        {
            try
            {
                ProductionHistoryLoginCheck();

                string SkyMarsServerIp = "", SkyMarsServerIpPort = "", UploadType = "";

                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetSkyMarsMachine(MachineId);
                jsonResponse = JObject.Parse(dataRequest);
                foreach (var item in jsonResponse["data"])
                {
                    SkyMarsServerIp = item["SkyMarsServerIp"].ToString();
                    SkyMarsServerIpPort = item["SkyMarsServerIpPort"].ToString();
                    UploadType = item["UploadType"].ToString();
                }
                InterfaceLib.IMsg iRemoting = null;
                StructMsg.nc_code _nc_code = new StructMsg.nc_code();
                StructMsg.NcName _nc_name = new StructMsg.NcName();
                StructMsg.Pwd _Pwd = new StructMsg.Pwd();
                if (ChannelServices.RegisteredChannels.Length == 0)
                    ChannelServices.RegisterChannel(new TcpChannel(), false);
                iRemoting = (IMsg)Activator.GetObject(typeof(IMsg), "tcp://" + SkyMarsServerIp + ":" + SkyMarsServerIpPort + "/RemoteObjectURI" + SkyMarsServerIpPort + "");

                string NC_PATH = "", O_NAME = "";
                //1:記憶體 ; 2:DataServer
                if (UploadType == "1")
                {
                    string[] FilePathArray = FilePath.Split('#');//第4次處理NC路徑 
                    NC_PATH = FilePathArray[0];//最終路徑
                    O_NAME = FilePathArray[1];//機台上檔案名稱

                    _Pwd.ConnectionKey = "pmc";
                    _Pwd.WritePwd = "pmc";
                    _nc_name.Name = O_NAME;

                    short DelRet = iRemoting.DEL_nc_mem(_Pwd, _nc_name);//執行刪除API
                    if (DelRet != 0)
                    {
                        logger.Error("操作者:" + UserNo + "，刪除NC程式失敗,錯誤代碼:" + DelRet);
                    }
                    else
                    {
                        _Pwd.ConnectionKey = "pmc";
                        _Pwd.WritePwd = "pmc";
                        _nc_name.Name = O_NAME;
                        _nc_code.NcCode = System.IO.File.ReadAllText(NC_PATH);
                        short UploadRet = iRemoting.UPLOAD_nc_mem(_Pwd, _nc_code);//執行上拋API
                        if (UploadRet != 0)
                        {
                            logger.Error("操作者:" + UserNo + "，上傳NC程式失敗,錯誤代碼:" + UploadRet);
                        }
                    }

                }
                else if (UploadType == "2")
                {
                    string[] FilePathArray = FilePath.Split('#');//第4次處理NC路徑 
                    NC_PATH = FilePathArray[0];//最終路徑
                    O_NAME = FilePathArray[1];//機台上檔案名稱

                    _Pwd.ConnectionKey = "pmc";
                    _Pwd.WritePwd = "pmc";
                    _nc_name.Name = O_NAME;
                    short DelRetFtp = iRemoting.DEL_nc_ftp(_Pwd, _nc_name);//執行刪除API
                    if (DelRetFtp != 0)
                    {
                        logger.Error("操作者:" + UserNo + "，刪除NC程式失敗,錯誤代碼:" + DelRetFtp);
                    }
                    else
                    {
                        _Pwd.ConnectionKey = "pmc";
                        _Pwd.WritePwd = "pmc";
                        _nc_name.Name = O_NAME;
                        _nc_code.NcCode = System.IO.File.ReadAllText(NC_PATH);
                        short UploadRetFtp = iRemoting.UPLOAD_nc_ftp(_Pwd, _nc_code);//執行上拋API
                        if (UploadRetFtp != 0)
                        {
                            logger.Error("操作者:" + UserNo + "，上傳NC程式失敗,錯誤代碼:" + UploadRetFtp);
                        }
                    }
                }
                else
                {
                    logger.Error("操作者:" + UserNo + "，異常回報:此機台未導入SkyMars");
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
        }
        #endregion

        #region  //NC加工程式上拋
        public string UPLOAD_NC_CODE(int MACHINE_ID, string NC_PATH)
        {
            string result = "", error = "";

            string SkyMarsServerIp = "", SkyMarsServerIpPort = "", UploadType = "";

            productionHistoryDA = new ProductionHistoryDA();
            dataRequest = productionHistoryDA.GetSkyMarsMachine(MACHINE_ID);
            jsonResponse = JObject.Parse(dataRequest);
            foreach (var item in jsonResponse["data"])
            {
                SkyMarsServerIp = item["SkyMarsServerIp"].ToString();
                SkyMarsServerIpPort = item["SkyMarsServerIpPort"].ToString();
                UploadType = item["UploadType"].ToString();
            }

            //NC_PATH = NC_PATH.Replace(" ", "");
            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                Arguments = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\"", UploadType, SkyMarsServerIp, SkyMarsServerIpPort, NC_PATH),
                CreateNoWindow = true,
                FileName = @"C:\WIN\Upload_NcCode\Upload_NcCode\bin\Debug\Upload_NcCode.exe",               
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false
            };

            using (Process process = Process.Start(processStartInfo))
            {
                using (StreamReader readerOutput = process.StandardOutput)
                {
                    result = readerOutput.ReadToEnd().Replace("\r", "").Replace("\n", "");
                    logger.Error(result.ToString());

                    if (result.Equals("OK"))
                    {
                        logger.Info("| UPLOAD_NC_CODE |" + result.ToString());
                    }
                    else
                    {
                        using (StreamReader readerError = process.StandardError)
                        {
                            error = readerError.ReadToEnd();
                            logger.Error("| UPLOAD_NC_CODE-ERROR |" + error.ToString() + "|Data==>MACHINE_ID:" + MACHINE_ID + ";SKM_IP:" + SkyMarsServerIp + ";SKYMAR_MACHINE:" + SkyMarsServerIpPort + ";NC_PATH:" + NC_PATH);
                        }
                    }
                    if (error.Length > 0) result += error;
                }
            }
            return result;
        }
        #endregion

        #region//自動手臂RobotArm
        #region//YCM
        #region //GetJigData -- 取得治具資訊
        [HttpPost]
        [Route("api/RobotArm/RobotArmGetJigData")]
        public void GetJigData(string Company = "", string SecretKey = "",
            string JigNo = "", string WorkingFlag = "N")
        {
            JObject jsonResponseNew = new JObject();
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RobotArmGetJigData");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetJigData(Company, JigNo, WorkingFlag);
                #endregion

                #region //Response
                if (JObject.Parse(dataRequest)["status"].ToString() == "error")
                {
                    jsonResponseNew = JObject.FromObject(new
                    {
                        status = "error",
                        msg = JObject.Parse(dataRequest)["msg"].ToString()
                    });
                }
                else if (JObject.Parse(dataRequest)["status"].ToString() == "errorForDA")
                {
                    jsonResponseNew = JObject.FromObject(new
                    {
                        status = "errorForDA",
                        msg = JObject.Parse(dataRequest)["msg"].ToString()
                    });
                }
                else
                {
                    if (JObject.Parse(dataRequest) != null)
                    {
                        jsonResponseNew = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "OK",
                            MoProcessResult = JObject.Parse(dataRequest)["data"],
                            NcDataResult = JObject.Parse(dataRequest)["data2"],
                        });
                    }
                    else
                    {
                        jsonResponseNew = JObject.FromObject(new
                        {
                            status = "success",
                            msg = JObject.Parse(dataRequest)["msg"].ToString()
                        });
                    }
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

            Response.Write(jsonResponseNew.ToString());
        }
        #endregion

        #region//YCM過站API
        [HttpPost]
        [Route("api/RobotArm/RobotArmTxJigBarcodeProcess")]
        public void TxJigBarcodeProcess(string Company = "", string SecretKey = "",
            string JigNo = "", int MoProcessId=-1, string MachineNo = "", int UserId = -1,
            string QtyBarcodeStatus = "", string QtyLotNgCauseNo = "", int QtyNextMoProcessId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RobotArmTxJigBarcodeProcess");
                #endregion

                #region//取得機台ID
                dataRequest = productionHistoryDA.GetMachineNo(MachineNo);
                JObject MachineNoJson = JObject.Parse(dataRequest);
                int MachineId = int.Parse(MachineNoJson["data"][0]["MachineId"].ToString());
                #endregion

                #region //Request
                dataRequest = productionHistoryDA.GetJigBarcodeData(JigNo);
                JObject ResultJson = JObject.Parse(dataRequest);
                string JigBarcodeIdData = ResultJson["data"][0]["JigBarcodeIdData"].ToString();                
                if (dataRequest.IndexOf("success") != -1) {
                    dataRequest = productionHistoryDA.UpdateJigBarcodeWorkingFlag(JigBarcodeIdData);
                }

                if (dataRequest.IndexOf("success") != -1)
                {
                    string BarcodeStatus = "";
                    string LotNgCauseNo = "";
                    int NextMoProcessId = -1;
                    dataRequest = productionHistoryDA.TxBarcodeProcess(JigNo, MoProcessId, MachineId, UserId, BarcodeStatus, LotNgCauseNo, NextMoProcessId, Company, "");
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

        #endregion

        #endregion

        #region //Keyence批量掃描API
        public string returnMsg = "";
        public string keyenceIP = "";
        public int keyencePort = -1;
        [HttpPost]
        public void KeyenceScan(int KeyenceId = -1)
        {
            try
            {
                if (KeyenceId <= 0) throw new SystemException("KeyenceId不能為空!!");

                productionHistoryDA = new ProductionHistoryDA();

                #region //確認Keyencey資料是否正確
                dataRequest = productionHistoryDA.GetKeyence(KeyenceId, "A");

                var dataRequestJson = JObject.Parse(dataRequest);

                if (dataRequest.IndexOf("success") != -1)
                {
                    if (dataRequestJson["data"].Count() <= 0 || dataRequestJson["data"].Count() > 1) throw new SystemException("Keyence資料錯誤!!");

                    keyenceIP = dataRequestJson["data"][0]["KeyenceIP"].ToString();
                    keyencePort = Convert.ToInt32(dataRequestJson["data"][0]["KeyencePort"]);
                }
                else
                {
                    throw new SystemException(dataRequestJson["msg"].ToString());
                }
                #endregion

                string message = "4C4F4E0D";
                byte[] data = HexStringToByteArray(message);
                // 创建TCP客户端套接字
                TcpClient clientSocket = new TcpClient();
                clientSocket.Connect(keyenceIP, keyencePort);
                if (clientSocket.Connected)
                {
                    NetworkStream stream = clientSocket.GetStream();
                    stream.Write(data, 0, data.Length);
                    // 接收服务器回复
                    byte[] responseData = new byte[256];
                    StringBuilder responseMessage = new StringBuilder();
                    int bytes = stream.Read(responseData, 0, responseData.Length);
                    responseMessage.Append(Encoding.ASCII.GetString(responseData, 0, bytes));

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = responseMessage
                    });

                    // 关闭套接字
                    stream.Dispose();
                    stream.Close();
                    clientSocket.Close();
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

        private static byte[] HexStringToByteArray(string hexString)
        {
            int len = hexString.Length;
            byte[] data = new byte[len / 2];
            for (int i = 0; i < len; i += 2)
            {
                data[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }
            return data;
        }

        public static bool IsConnet = true;
        public void Connet(string Iptxt, int Port)//接收参数是目标ip地址和目标端口号。客户端无须关心本地端口号
        {
            //创建一个新的Socket对象
            Socket client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IsConnet = true;//注意，此处是全局变量，将其设置为true
                            //将方法写进线程中
            Thread thread = new Thread(() =>
            {
                while (IsConnet)//循环
                {
                    try
                    {
                        client.Connect(IPAddress.Parse(Iptxt), Port);//尝试连接，失败则会跳去catch
                        IsConnet = false;//成功连接后修改bool值为false,这样下一步循环就不再执行。
                        break;//在此处加上break，成功就跳出循环，避免死循环
                    }
                    catch
                    {
                        client.Close();//先关闭
                                       /*使用新的客户端资源覆盖，上一个已经废弃。如果继续使用以前的资源进行连接，
                                       即使参数正确， 服务器全部打开也会无法连接*/
                        client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                        Thread.Sleep(1000);//等待1s再去重连
                    }
                }
                /*这里不一样就是放接收线程，在连接上后break出来，执行。
                因为需要带参数，所以要用到特别的ParameterizedThreadStart，
                然后开始线程。↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓*/
                Thread thread2 = new Thread(new ParameterizedThreadStart(ClientReceiveData));//接收线程方法
                thread2.IsBackground = true;//该值指示某个线程是否为后台线程。
                thread2.Start(client);//参数是用我们自建的Socket对象，就是上面的Socket client=new……

            });
            thread.IsBackground = true;//设置为后台线程，在程序退出时自己会自动释放
            thread.Start();//开始执行线程
        }

        public void ClientReceiveData(object socket)//TCPClient消息的方法
        {
            var ProxSocket = socket as Socket;//处理上一步传过来的Socket函数
            byte[] data = new byte[1024 * 1024];//接收消息的缓冲区
            while (!IsConnet)//同样循环中止的条件
            {
                int len = 0;//记录消息长度，以及判断是否连接
                try
                {
                    //连接函数Receive会将数据放入data,从0开始放，之后返回数据长度。
                    len = ProxSocket.Receive(data, 0, data.Length, SocketFlags.None);
                }
                catch (Exception)
                {
                    //异常退出
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    Connet(keyenceIP, keyencePort);
                    IsConnet = false;
                    return;
                }

                if (len <= 0)
                {
                    ProxSocket.Shutdown(SocketShutdown.Both);//中止传输
                    ProxSocket.Close();//关闭
                    Connet(keyenceIP, keyencePort);
                    IsConnet = false;
                    return;
                }

                //这里做你想要对消息做的处理
                returnMsg = Encoding.Default.GetString(data, 0, len);//二进制数组转换成字符串……
            }
        }
        #endregion

        #region //BarcodeProcessByRoboticArm MTF機械手臂過站API
        [HttpPost]
        [Route("api/MES/BarcodeProcessByRoboticArm")]
        public void BarcodeProcessByRoboticArm(string Company = "", string SecretKey = "", string BarcodeNo = "", int MoProcessId = -1, string UserNo = "", int MachineId = -1, int ProcessCount = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "BarcodeProcessByRoboticArm");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.BarcodeProcessByRoboticArm(Company, BarcodeNo, MoProcessId, UserNo, MachineId, ProcessCount);
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

        #region //NgBarcodeProcessByRoboticArm MTF機械手臂刷取NG條碼
        [HttpPost]
        [Route("api/MES/NgBarcodeProcessByRoboticArm")]
        public void NgBarcodeProcessByRoboticArm(string Company = "", string SecretKey = "", string UserNo = "", string BarcodeNo = "", string NgBarcodeNo = "", int NgBarcodeQty = -1, string CauseId = "", int MachineId = -1, string CauseNo = "", int MoProcessId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "NgBarcodeProcessByRoboticArm");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddLotNgBarcodeProcessByRoboticArm(Company, UserNo, BarcodeNo, NgBarcodeNo, NgBarcodeQty, CauseId, MachineId, CauseNo, MoProcessId);
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

        #region //GetMoProcessByRoboticArm MTF機械手臂取得製令製程資料
        [HttpPost]
        [Route("api/MES/GetMoProcessByRoboticArm")]
        public void GetMoProcessByRoboticArm(string Company = "", string SecretKey = "", string WoErpFullNo = "", int MachineId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetMoProcessByRoboticArm");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessByRoboticArm(Company, WoErpFullNo, MachineId);
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

        #region//GetMachineSchedule 查詢機台排程
        [HttpPost]
        public void GetMachineSchedule(int MachineId = -1)
        {
            using (HttpClient client = new HttpClient())
            {
                string apiUrl = $"http://192.168.20.46:2536/APS_one/GetMachineSchedule?MachineId={MachineId}";
                client.Timeout = TimeSpan.FromMinutes(60);
                using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, apiUrl))
                {
                    using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                    {
                        if (httpResponseMessage.IsSuccessStatusCode)
                        {
                            dataRequest = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                            dataRequest = dataRequest.Replace("result", "data");
                            if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                            {                               
                                jsonResponse = BaseHelper.DAResponse(dataRequest);                               
                                Response.Write(jsonResponse.ToString());
                            }
                            else
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "error",
                                    msg = JObject.Parse(dataRequest)["message"].ToString()
                                });
                                #endregion
                            }
                        }
                        else
                        {
                            throw new SystemException("伺服器連線異常");
                        }
                    }
                }
            }
        }
        #endregion

        #region//UpdateTempMtfQcMeasureData 解析暫存的MTF數據 -- Luca 2024-09-09
        [HttpPost]
        [Route("api/MES/apiTempMtfQcMeasureData")]
        public void UpdateTempQcMeasureData(string Company = "", string SecretKey = "",string RunStatus = "N")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateTempMtfQcMeasureData");
                #endregion

                #region //Request
                dataRequest = productionHistoryDA.UpdateTempMtfQcMeasureData(Company,RunStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddBarcodeComponentByArm 手臂上料
        [HttpPost]
        [Route("api/MES/AddBarcodeComponentByArm")]
        public void AddBarcodeComponentByArm(string Company = "", string SecretKey = "", string UserNo = "", string BarcodeNo  ="", string ComponentNo = "", int MoProcessId = -1, string ComponentMtlitemNo  ="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "AddBarcodeComponentByArm");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeComponentByArm(Company, UserNo, BarcodeNo, ComponentNo, MoProcessId, ComponentMtlitemNo);
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

        #region //BarcodeMergeByArm 手臂拆併盤
        [HttpPost]
        [Route("api/MES/BarcodeMergeByArm")]
        public void BarcodeMergeByArm(string Company = "", string SecretKey = "", string UserNo = "", int MoProcessId = -1 , string BarcodeNo = "", string OKTrayNo = "", string MTFBTrayNo = "", int MergeQty = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "BarcodeMergeByArm");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.BarcodeMergeByArm(Company, UserNo, MoProcessId, BarcodeNo, OKTrayNo, MTFBTrayNo, MergeQty);
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



        #region //晶彩 檢挑補機台 【API01】 取得製令製程資料
        [HttpPost]
        [Route("api/MES/apiGetMoProcess")]
        public void ApiGetMoProcess(string Company = "", string SecretKey = "", string WoErpFullNo = "", int MachineId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetMoProcess");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMoProcessByRoboticArm(Company, WoErpFullNo, MachineId);
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

        #region//晶彩 檢挑補機台 【API02】製令開工
        [HttpPost]
        [Route("api/MES/apiBarcodeProcess")]
        public void ApiBarcodeProcess(string Company = "", string SecretKey = "", string BarcodeNo = "", int MoProcessId = -1, string UserNo = "", int MachineId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "BarcodeProcess");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.BarcodeProcessByRoboticArm(Company, BarcodeNo, MoProcessId, UserNo, MachineId, -1);
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

        #region//晶彩 檢挑補機台 【API03】不良原因與數量紀錄
        [HttpPost]
        [Route("api/MES/apiNgBarcodeProcess")]
        public void ApiNgBarcodeProcess(string Company = "", string SecretKey = "", string UserNo = "", string BarcodeNo = "", string NgBarcodeNo = "", int NgBarcodeQty = -1, string CauseId = "", int MachineId = -1, string CauseNo = "", int MoProcessId = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "NgBarcodeProcess");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddLotNgBarcodeProcessByRoboticArm(Company, UserNo, BarcodeNo, NgBarcodeNo, NgBarcodeQty, CauseId, MachineId, CauseNo, MoProcessId);
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

        #region//晶彩 檢挑補機台 【API04】拆並盤紀錄
        [HttpPost]
        [Route("api/MES/apiLotBarcodeQtyMerge")]
        public void ApiLotBarcodeQtyMerge(string Company = "", string SecretKey = "", string UserNo = "", int MoId = -1, string FromBarcodeNo = "", string ToBarcodeNo = "", int TransactionQty = -1, string MinBarcodes = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "LotBarcodeQtyMerge");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddLotBarcodeQtyMergeForApi(Company, UserNo, MoId, FromBarcodeNo, ToBarcodeNo, TransactionQty, MinBarcodes);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //晶彩 檢挑補機台 【API05】入庫後拆並盤紀錄
        [HttpPost]
        [Route("api/MES/apiLotBarcodeQtyMergeAfterInward")]
        public void ApiLotBarcodeQtyMergeAfterInward(string Company, string SecretKey = "", string UserNo = "", string FromBarcodeNo = "", string ToBarcodeNo = "", int TransactionQty = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "LotBarcodeQtyMergeAfterInward");
                #endregion

                #region //Request
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.BarcodeMergerTransactionAfterInwardForApi(Company, UserNo, FromBarcodeNo, ToBarcodeNo, TransactionQty);
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

        #endregion

        #region //MODAL
        #region //資料夾
        public class Folder
        {
            public string FolderName { get; set; }
            public string FolderPath { get; set; }
        }
        #endregion

        #region //檔案
        public class FileInfoModal
        {
            public string FileName { get; set; }
            public string FilePath { get; set; }
            public string FolderPath { get; set; }
            public string WhitelistPath { get; set; }
            public string ListId { get; set; }
        }
        #endregion
        #endregion

        #region //Custom
        #region //登入
        [HttpPost]
        public void ProductionHistoryLogin(string UserNo, string SystemKey)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetSubSystemLogin(SystemKey, UserNo, "ProductionHistory");
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    LoginLog(Convert.ToInt32(result[0]["UserId"]), UserNo, "ProductionHistory");
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

        #region //View檢查登入
        [NonAction]
        public void ProductionHistoryLoginCheck()
        {
            bool verify = LoginVerify("ProductionHistory");

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "ProductionHistory");

                var uri = new Uri(redirectUrl);
                var newQueryString = HttpUtility.ParseQueryString(uri.Query);

                string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path).Replace("http://", "").Replace(Request.Url.Authority, "");
                redirectUrl = newQueryString.Count > 0 ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString) : pagePathWithoutQueryString;
                if (redirectUrl != "/") redirectLoginUrl += "?preUrl=" + Uri.EscapeDataString(redirectUrl);

                HttpContext.Response.Redirect(redirectLoginUrl);
            }
        }
        #endregion

        #region //Api檢查登入
        [NonAction]
        public void ProductionHistoryApiLoginCheck()
        {
            bool verify = LoginVerify("ProductionHistory");

            if (!verify)
            {
                throw new SystemException("系統已登出，請重新登入");
            }
        }
        #endregion
        #endregion

        #region //Tray相關

        #region //Get

        #region //GetTrayBarcodeInfo -- 取得TRAY盤條碼資訊
        [HttpPost]
        public void GetTrayBarcodeInfo(int MoId = -1, string TrayNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetTrayBarcodeInfo(MoId, TrayNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //GetTrayBindLens -- 取得TRAY盤已綁定鏡頭條碼列表 
        [HttpPost]
        public void GetTrayBindLens(int ParentBarcodeId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetTrayBindLens(ParentBarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //GetNgTrayInfo //取得Ng Tray盤資訊
        [HttpPost]
        public void GetNgTrayInfo(string TrayNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetNgTrayInfo(TrayNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region  GetTrayMoBom//取得Tray盤物料
        [HttpPost]
        public void GetTrayMoBom(int MtlItemId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetTrayMoBom(MtlItemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region  GetNgLensBarcodeInfo//取得Lens物料
        [HttpPost]
        public void GetNgLensBarcodeInfo(string BarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetNgLensBarcodeInfo(BarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region  GetJudgeLensTray //判別為LENS/TRAY
        [HttpPost]
        public void GetJudgeLensTray(string inputNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetJudgeLensTray(inputNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region  //GetMachineMoInfo --取得機台製令資訊
        [HttpPost]
        public void GetMachineMoInfo(int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetMachineMoInfo(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region  //GetLotNumber --取得批號條碼資料
        [HttpPost]
        public void GetLotNumber(string LnBarcodeNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetLotNumber(LnBarcodeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //GetTrayLOTBarcodeInfo -- 取得TRAY盤條碼資訊
        [HttpPost]
        public void GetTrayLOTBarcodeInfo(int MoId = -1, string inputNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetTrayLOTBarcodeInfo(MoId, inputNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //GetBarcodeTrayKence -- 取得條碼資料 -
        [HttpPost]
        public void GetBarcodeTrayKence(string TrayNo = "")
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.GetBarcodeTrayKence(TrayNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //ADD

        #region //AddBarcodeToBindLens --Tray綁鏡頭條碼

        [HttpPost]
        public void AddBarcodeToBindLens(string TrayNo = "", string XTrayLocation = "", string YTrayLocation = "", string BarcodeNo = "", int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeToBindLens(TrayNo, XTrayLocation, YTrayLocation, BarcodeNo, MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //AddBarcodeToBindLensLot --Tray綁鏡頭條碼

        [HttpPost]
        public void AddBarcodeToBindLensLot(string LotNo = "", string XTrayLocation = "", string YTrayLocation = "", string BarcodeNo = "", int MoId = -1, int MoProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.AddBarcodeToBindLensLot(LotNo, XTrayLocation, YTrayLocation, BarcodeNo, MoId, MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateTrayBindtTurnover -- 周轉Tray盤與鏡頭條碼移轉
        [HttpPost]
        public void UpdateTrayBindtTurnover(string oldTrayNo = "", string newTrayNo = "", int moProcessId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateTrayBindtTurnover(oldTrayNo, newTrayNo, moProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateTrayNg -- NGTRAY物料
        [HttpPost]
        public void UpdateTrayNg(string inputNo  ="", string NgBom = "", int NgcountQty = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateTrayNg(inputNo, NgBom, NgcountQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateLensNg -- NG LENS物料
        [HttpPost]
        public void UpdateLensNg(string inputNo = "", string NgBom = "", int NgcountQty = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateLensNg(inputNo, NgBom, NgcountQty);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateMachineConsume --更換耗材
        [HttpPost]
        public void UpdateMachineConsume(int MachineId = -1, string LnBarcodeNo = "", int BarcodeQty = -1, string OpeningDate = "", int Exps = -1, int AllowQty = -1, int LotNumberId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateMachineConsume(MachineId, LnBarcodeNo, BarcodeQty, OpeningDate, Exps, AllowQty, LotNumberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateUnbindMachineConsume --更換耗材
        [HttpPost]
        public void UpdateUnbindMachineConsume(int MachineId = -1)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UpdateUnbindMachineConsume(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //UddateBindingBarcode -- 解除綁定條碼資料(TRAY盤條碼用)

        [HttpPost]
        public void UddateBindingBarcode(int BarcodeId)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UddateBindingBarcode(BarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UddateAllBindingBarcode -- 解除綁定條碼資料(TRAY盤條碼用)

        [HttpPost]
        public void UddateAllBindingBarcode(int BarcodeId)
        {
            try
            {
                ProductionHistoryLoginCheck();

                #region //Request       
                productionHistoryDA = new ProductionHistoryDA();
                dataRequest = productionHistoryDA.UddateAllBindingBarcode(BarcodeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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