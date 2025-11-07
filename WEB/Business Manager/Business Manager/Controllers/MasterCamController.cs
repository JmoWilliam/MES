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
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

#region
#endregion

namespace Business_Manager.Controllers
{
    public class MasterCamController : WebController
    {
        private MasterCamDA masterCamDA = new MasterCamDA();
        string CadFolderPath = @"";
        public void GetCadFloderPath(int CompanyId)
        {
            switch (CompanyId)
            {
                case 2:
                    CadFolderPath = @"~/MasterCam";
                    break;
                case 4:
                    CadFolderPath = @"~/MasterCam";
                    break;
                case 3:
                    CadFolderPath = @"~/MasterCam";
                    break;
                default:
                    throw new SystemException("此公司別尚未維護設計圖上傳路徑!!");
            }
        }

        #region //View
        // GET: MasterCam
        public ActionResult CncProgramSetting()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult CncProgramLog()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MillProgramLog() //銑床報工清單
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult MillMachineTool() //銑床機台刀具維護頁面
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get

        #region//取得MES製令設計圖資料
        [HttpPost]
        public void GetMoRdDesign(int MoId = -1,string OtherInfo=""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read,constrained-data");

                #region //Request 
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMoRdDesign(MoId, OtherInfo, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMoProcess --取得製令製程
        [HttpPost]
        public void GetMoProcess(int MoId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request       
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMoProcess(MoId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncProgram --取得編程
        [HttpPost]
        public void GetCncProgram(int MoProcessId = -1,int CncProgId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request   
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncProgram(MoProcessId, CncProgId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetProcessMachine --取得機台
        [HttpPost]
        public void GetProcessMachine(int MoProcessId = -1,int MachineId=-1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request    
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetProcessMachine(MoProcessId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncParameter --取得編程參數
        [HttpPost]
        public void GetCncParameter(int CncProgId = -1,string CncParamStatus="")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request 
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncParameter(CncProgId, CncParamStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//GetCncParameterOption --取得參數選項
        [HttpPost]
        public void GetCncParameterOption(int CncParamId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request   
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncParameterOption(CncParamId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region//GetToolMachineSetting --取得機台所需刀具參數項目設定
        [HttpPost]
        public void GetToolMachineSetting(int ToolId = -1, string ToolNo = "")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request  
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetToolMachineSetting(ToolId, ToolNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetProcessMachineToolCount --取得製程機台所綁工具上限數
        [HttpPost]
        public void GetProcessMachineToolCount(int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request  
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetProcessMachineToolCount(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetToolMachine --取得機台上工具綁定紀錄
        [HttpPost]
        public void GetToolMachine(int MachineId = -1,int ProcessId=-1,int ModeId=-1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request   
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetToolMachine(MachineId, ProcessId, ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetToolMachineParameter --取得機台上工具的設定值
        [HttpPost]
        public void GetToolMachineParameter(int ToolMachineId = -1,int ToolId=-1, int ProcessId = -1, int ModeId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request 
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetToolMachineParameter(ToolMachineId,ToolId,ProcessId,ModeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncProgramJmoFile --取得編程JMO檔案
        [HttpPost]
        public void GetCncProgramJmoFile(int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request  
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncProgramJmoFile(FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncProgramRequest --取得CncProgramRequest參數
        [HttpPost]
        public void GetCncProgramRequest(int CncProgLogId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncProgramRequest(CncProgLogId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncProgramResponsest --取得CncProgramResponsest結果
        [HttpPost]
        public void GetCncProgramResponsest(int CncProgLogId = -1,string CncParamNo="")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request 
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncProgramResponsest(CncProgLogId, CncParamNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncProgramWork --取得CncProgramWork結果
        [HttpPost]
        public void GetCncProgramWork(int CncProgramWorkId=-1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncProgramWork(CncProgramWorkId,OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetCncProgramWorkFiles --取得報工上傳檔案結果
        [HttpPost]
        public void GetCncProgramWorkFiles(int CncProgramWorkId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request   
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetCncProgramWorkFiles(CncProgramWorkId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMoProcessParameter --取得生產模式與製程
        [HttpPost]
        public void GetMoProcessParameter(int MoProcessId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMoProcessParameter(MoProcessId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //GetParameterHistoryData -- 取得編程後參數數值
        [HttpPost]
        public void GetParameterHistoryData(int CncProgramWorkId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetParameterHistoryData(CncProgramWorkId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region// GetMaxToolBlockSeq --取得銑床刀具現況最大刀座編號
        [HttpPost]
        public void GetMaxToolBlockSeq(int MachineId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMaxToolBlockSeq(MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMillToolSpec 取得銑床的刀具規格
        [HttpPost]
        public void GetMillToolSpec(string ToolNo = "")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMillToolSpec(ToolNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMillMachineTool 取得銑床機台刀具資訊
        [HttpPost]
        public void GetMillMachineTool(string OrderBy = "", int PageIndex = -1, int PageSize = -1, int MachineId = -1,string MachineNo="",int MillMachineToolId=-1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMillMachineTool(OrderBy, PageIndex, PageSize,MachineId, MachineNo, MillMachineToolId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMillProgramWork 銑床報工明細
        [HttpPost]
        public void GetMillProgramWork(string OrderBy = "", int PageIndex = -1, int PageSize = -1
            ,int MoId=-1, string WoErpFullNo = "",int MoProcessId=-1,int MachineId = -1,string StartDateTime = "", string EndDateTime = "")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMillProgramWork(OrderBy, PageIndex, PageSize
                    , MoId, WoErpFullNo, MoProcessId, MachineId, StartDateTime, EndDateTime);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMillData 銑床檔案下載
        [HttpPost]
        public void GetMillData(string OrderBy = "", int PageIndex = -1, int PageSize = -1
                    , string DataType = "", int MillProgramWorkId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "read");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.GetMillData(OrderBy, PageIndex, PageSize
                    , DataType, MillProgramWorkId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region//編程程式基本資料
        [HttpPost]
        public void AddCncProgram(int ProcessId, string CncProgName, string CncProgApi)
        {
            try
            {
                WebApiLoginCheck("CncProgramSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncProgram(ProcessId, CncProgName, CncProgApi);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//編程程式檔案基本資料
        [HttpPost]
        public void AddCncProgramFile(int CncProgId, string FileDesc, int FileId)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncProgramFile(CncProgId, FileDesc, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//編程程式參數基本資料
        [HttpPost]
        public void AddCncParameter(int CncProgId, string CncParamNo, string CncParamName, string CncParamType, string CncParamOption
        , string CncParamUomId, string CncParamValue, string CncParamStatus)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncParameter(CncProgId, CncParamNo, CncParamName, CncParamType, CncParamOption
                    , CncParamUomId, CncParamValue, CncParamStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//編程程式刀具基本資料
        [HttpPost]
        public void AddCncProgramKnife(int CncProgId, string KnifeName, string KnifeKey)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncProgramKnife(CncProgId, KnifeName, KnifeKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddCncProgramLog --編程執行紀錄資料 新增 --Ding 2022.10.17
        [HttpPost]
        public void AddCncProgramLog(int CncProgId, int MoProcessId,int MachineId)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncProgramLog(CncProgId, MoProcessId, MachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddCncProgramRequest --編程需求參數數值 新增 --Ding 2022.11.7
        [HttpPost]
        public void AddCncProgramRequest(int CncProgLogId, string RequestKey,string RequestValues)
        {
            try
            {                
                WebApiLoginCheck("CncProgramLog", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncProgramRequest(CncProgLogId, RequestKey, RequestValues);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion
                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddToolMachineParameter --機台所需刀具參數 新增
        [HttpPost]
        public void AddToolMachineParameter(string ToolParameterData = "")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "add");
                masterCamDA = new MasterCamDA();

                int MachineId = -1,MachineLocation = -1,ToolClassIdSettingId = -1;
                string ToolNo = "", ParameterValue = "";
                if(ToolParameterData == "") throw new SystemException("【機台所需刀具參數JSON】不能為空!");
                JObject ToolMachineParameterJson = JObject.Parse(ToolParameterData);
                int ToolMachineId = -1;
                for (int i = 0; i < ToolMachineParameterJson["toolParameterInfo"].Count(); i++)
                {
                    MachineId = int.Parse(ToolMachineParameterJson["toolParameterInfo"][i]["MachineId"].ToString());
                    MachineLocation = int.Parse(ToolMachineParameterJson["toolParameterInfo"][i]["MachineLocation"].ToString());
                    ToolClassIdSettingId = int.Parse(ToolMachineParameterJson["toolParameterInfo"][i]["ToolClassIdSettingId"].ToString());
                    ToolNo = ToolMachineParameterJson["toolParameterInfo"][i]["ToolNo"].ToString();
                    ParameterValue = ToolMachineParameterJson["toolParameterInfo"][i]["ParameterValue"].ToString();

                    int AddCount = i;
                    #region //Request
                    dataRequest = masterCamDA.AddToolMachineParameter(MachineId, ToolNo, MachineLocation, ToolClassIdSettingId, ParameterValue, AddCount, ToolMachineId);

                    var dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequest.IndexOf("errorForDA") != -1)
                    {
                        throw new SystemException(dataRequestJson["msg"].ToString());
                    }
                    else
                    {
                        foreach (var item in dataRequestJson)
                        {
                            if (item.Key == "data")
                            {
                                foreach (var item2 in item.Value)
                                {
                                    ToolMachineId = Convert.ToInt32(item2["ToolMachineId"]);
                                }
                            }
                        }
                    }
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

        #region//AddCncProgramResponsest --編程回傳數值 新增
        [HttpPost]
        public void AddCncProgramResponsest(int CncProgLogId=-1,int CncProgId = -1, int MachineId = -1,int MoProcessId=-1)
        {
            string programPath = "";
            int CncProgNo = -1, CncMachineNo=-1;
            try
            {
                WebApiLoginCheck("CncProgramLog", "add");
                masterCamDA = new MasterCamDA();

                #region //查詢編程路徑                               
                dataRequest = masterCamDA.GetCncProgram(-1, CncProgId);                
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                programPath = jsonResponse["result"][0]["CncProgApi"].ToString();
                CncProgNo = int.Parse(jsonResponse["result"][0]["CncProgNo"].ToString());
                #endregion

                #region//查詢編程機台
                dataRequest = masterCamDA.GetProcessMachine(MoProcessId, MachineId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);               
                CncMachineNo = int.Parse(jsonResponse["result"][0]["CncMachineNo"].ToString());
                #endregion

                #region //查詢需求參數
                dataRequest = masterCamDA.GetCncProgramRequest(CncProgLogId);                
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                JObject NC_PARAMETER = new JObject();
                string NC_PARAMETER_NO = "", REQUEST_VALUES="";
                for (int i = 0; i < jsonResponse["result"].Count(); i++)
                {
                    NC_PARAMETER_NO = jsonResponse["result"][i]["CncParamNo"].ToString();                   
                    if (NC_PARAMETER_NO == "product_url") {
                        REQUEST_VALUES = jsonResponse["result"][i]["RequestValues"].ToString();
                        //判斷接收的Jmo是整數還是字串
                        int n;
                        bool isNumeric = int.TryParse(REQUEST_VALUES, out n);
                        if (isNumeric) {
                            REQUEST_VALUES = BaseToRealPath(int.Parse(REQUEST_VALUES));
                            JObject JmoPathJson = JObject.Parse(REQUEST_VALUES);
                            string JmoPath = JmoPathJson["file_url"].ToString();
                            NC_PARAMETER.Add(NC_PARAMETER_NO, JmoPath);
                        }
                        else {                            
                            string JmoPath = CopyNasFileToMes(REQUEST_VALUES);
                            NC_PARAMETER.Add(NC_PARAMETER_NO, JmoPath);
                        }

                    } else if (NC_PARAMETER_NO == "KNIFE_CONTEXT") {
                        REQUEST_VALUES = jsonResponse["result"][i]["RequestValues"].ToString();
                        int n;
                        bool isNumeric = int.TryParse(REQUEST_VALUES, out n);
                        if (!isNumeric) {                            
                            REQUEST_VALUES = jsonResponse["result"][i]["RequestValues"].ToString().Replace("{\"KNIFE_CONTEXT\":", "");
                            REQUEST_VALUES = REQUEST_VALUES.Substring(0, REQUEST_VALUES.Length - 1);
                            NC_PARAMETER.Add(NC_PARAMETER_NO, REQUEST_VALUES);
                        }                        
                    }else{                       
                        REQUEST_VALUES = jsonResponse["result"][i]["RequestValues"].ToString();
                        NC_PARAMETER.Add(NC_PARAMETER_NO, REQUEST_VALUES);
                    }
                    
                }
                NC_PARAMETER.Add("txtMachineID", CncMachineNo);
                NC_PARAMETER.Add("NC_PROCESS_LOG_ID", CncProgLogId);
                NC_PARAMETER.Add("NC_PROGRAM_ID", CncProgNo);

                #endregion

                #region//呼叫編程程式
                string ResponsestValues = "", CncParamNo="";
                string CallCncProgramPythonResult = NC_Python(NC_PARAMETER.ToString(Newtonsoft.Json.Formatting.None), programPath);//呼叫編程程式
                bool isContains = CallCncProgramPythonResult.Contains("Traceback (most recent call last):");
                if (isContains)
                {
                    string res = "{\"status\" : \"error\",\"msg\": \"error\",\"error_desc\":\"" + CallCncProgramPythonResult + "\"}";
                    logger.Error(res + "| NC_PROCESS_LOG_ID |" + CncProgLogId);
                }
                JObject CallCncProgramPythonResultJson = JObject.Parse(CallCncProgramPythonResult);
                if (CallCncProgramPythonResultJson["STATUS"].ToString() == "N")
                {
                    if (CallCncProgramPythonResultJson["message"] != null)
                    {
                        for (int i = 0; i < CallCncProgramPythonResultJson["message"].Count(); i++)
                        {
                            ResponsestValues = CallCncProgramPythonResultJson["message"][i].ToString();
                            CncParamNo = "message";                           
                            dataRequest = masterCamDA.AddCncProgramResponsest(CncProgLogId, CncParamNo, ResponsestValues, -1); //新增回傳資料                            
                        }
                    }
                }
                else {
                    if (CallCncProgramPythonResultJson["message"] != null)
                    {
                        for (int i = 0; i < CallCncProgramPythonResultJson["message"].Count(); i++)
                        {
                            ResponsestValues = CallCncProgramPythonResultJson["message"][i].ToString();
                            CncParamNo = "message";
                            dataRequest = masterCamDA.AddCncProgramResponsest(CncProgLogId, CncParamNo, ResponsestValues, -1); //新增回傳資料
                        }
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

        #region//AddCncDirectWork --直接編程報工 新增
        [HttpPost]
        public void AddCncDirectWork(int CncProgLogId = -1, int MoProcessId = -1, int MachineId = -1, string StartWorkDate = "", string EndWorkDate = "",int WorkCT = -1,
            string FileArray="")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncDirectWork(CncProgLogId, MoProcessId, MachineId, StartWorkDate, EndWorkDate, WorkCT, FileArray);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion
                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddCncProgramWork --編程報工 新增
        [HttpPost]
        public void AddCncProgramWork(int CncProgLogId = -1, int MoProcessId = -1, int MachineId = -1, string StartWorkDate = "", string EndWorkDate = "", int WorkCT = -1,
            string FileArray = "")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddCncProgramWork(CncProgLogId, MoProcessId, MachineId, StartWorkDate, EndWorkDate, WorkCT, FileArray);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion
                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMillMachineTool --銑床機台刀具 新增
        [HttpPost]
        public void AddMillMachineTool(string ToolNo = "", int MachineId = -1, string ToolBlockSeq = "", string ToolBlockName = ""
            , string MillToolD = "", string MillToolRone = "",string MillToolB = "", string MillToolFL = "", string MillToolL = "", string MillToolHD = ""
            , string MillToolR = "", string MillToolRPM = "", string MillToolFPR = "")
        {
            try
            {
                WebApiLoginCheck("MillMachineTool", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.AddMillMachineTool(ToolNo, MachineId, ToolBlockSeq, ToolBlockName
                    , MillToolD, MillToolRone, MillToolB, MillToolFL, MillToolL, MillToolHD
                    , MillToolR, MillToolRPM, MillToolFPR);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//編程程式基本資料
        [HttpPost]
        public void UpdateCncProgram(int CncProgId, int ProcessId, string CncProgName, string CncProgApi)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "update");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.UpdateCncProgram(CncProgId, ProcessId, CncProgName, CncProgApi);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//編程程式檔案基本資料
        [HttpPost]
        public void UpdateCncProgramFile(int CncProgramFileId, int CncProgId, string FileDesc, int FileId)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.UpdateCncProgramFile(CncProgramFileId, CncProgId, FileDesc, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//編程程式參數基本資料
        [HttpPost]
        public void UpdateCncParameter(int CncParamId, int CncProgId, string CncParamNo, string CncParamName, string CncParamType, string CncParamOption
        , string CncParamUomId, string CncParamValue, string CncParamStatus)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.UpdateCncParameter(CncParamId, CncProgId, CncParamNo, CncParamName, CncParamType, CncParamOption
                    , CncParamUomId, CncParamValue, CncParamStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//編程程式刀具基本資料
        [HttpPost]
        public void UpdateCncProgramKnife(int CncProgramKnifeId, int CncProgId, string KnifeName, string KnifeKey)
        {
            try
            {
                WebApiLoginCheck("AutomaticProgrammingSetting", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.UpdateCncProgramKnife(CncProgramKnifeId, CncProgId, KnifeName, KnifeKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateCncProgramResponsest --編程回傳數值 修改
        [HttpPost]
        public void UpdateCncProgramResponsest(int CncProgResponsestId,string ResponsestValues, int FileId)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "update");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.UpdateCncProgramResponsest(CncProgResponsestId, ResponsestValues, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//UpdateToolMachineParameter --機台所需刀具參數 修改
        [HttpPost]
        public void UpdateToolMachineParameter(int ToolMachineId = -1, string ToolParameterData = "")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "update");
                masterCamDA = new MasterCamDA();

                int MachineId = -1, MachineLocation = -1, ToolClassIdSettingId = -1;
                string ToolNo = "", ParameterValue = "";
                if (ToolParameterData == "") throw new SystemException("【機台所需刀具參數JSON】不能為空!");
                JObject ToolMachineParameterJson = JObject.Parse(ToolParameterData);
                for (int i = 0; i < ToolMachineParameterJson["toolParameterInfo"].Count(); i++)
                {
                    MachineId = int.Parse(ToolMachineParameterJson["toolParameterInfo"][i]["MachineId"].ToString());
                    MachineLocation = int.Parse(ToolMachineParameterJson["toolParameterInfo"][i]["MachineLocation"].ToString());
                    ToolClassIdSettingId = int.Parse(ToolMachineParameterJson["toolParameterInfo"][i]["ToolClassIdSettingId"].ToString());
                    ToolNo = ToolMachineParameterJson["toolParameterInfo"][i]["ToolNo"].ToString();
                    ParameterValue = ToolMachineParameterJson["toolParameterInfo"][i]["ParameterValue"].ToString();

                    int AddCount = i;
                    #region //Request
                    dataRequest = masterCamDA.UpdateToolMachineParameter(ToolMachineId, ToolClassIdSettingId, ParameterValue, MachineId, MachineLocation, ToolNo, AddCount);

                    var dataRequestJson = JObject.Parse(dataRequest);
                    if (dataRequest.IndexOf("errorForDA") != -1)
                    {
                        throw new SystemException(dataRequestJson["msg"].ToString());
                    }
                    else
                    {
                        foreach (var item in dataRequestJson)
                        {
                            if (item.Key == "data")
                            {
                                foreach (var item2 in item.Value)
                                {
                                    ToolMachineId = Convert.ToInt32(item2["ToolMachineId"]);
                                }
                            }
                        }
                    }
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

        #region //UpdateCncProgramWork --報工紀錄 修改
        [HttpPost]
        public void UpdateCncProgramWork(int CncProgramWorkId=-1, string StartWorkDate = "", string EndWorkDate = "", string WorkCT="")
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "update");
                masterCamDA = new MasterCamDA();

                #region //Request
                dataRequest = masterCamDA.UpdateCncProgramWork(CncProgramWorkId, StartWorkDate, EndWorkDate, WorkCT);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMillMachineTool --銑床機台刀具 修改
        [HttpPost]
        public void UpdateMillMachineTool(int MillMachineToolId=-1,string ToolNo = "", int MachineId = -1, string ToolBlockSeq = "", string ToolBlockName = ""
            , string MillToolD = "", string MillToolRone = "", string MillToolB = "", string MillToolFL = "", string MillToolL = "", string MillToolHD = ""
            , string MillToolR = "", string MillToolRPM = "", string MillToolFPR = "",int OriMachineId=-1,string OriToolNo = "")
        {
            try
            {
                WebApiLoginCheck("MillMachineTool", "update");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.UpdateMillMachineTool(MillMachineToolId, ToolNo, MachineId, ToolBlockSeq, ToolBlockName
                    , MillToolD, MillToolRone, MillToolB, MillToolFL, MillToolL, MillToolHD
                    , MillToolR, MillToolRPM, MillToolFPR, OriMachineId, OriToolNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region//DeleteCncProgramWork --編程報工紀錄 刪除
        [HttpPost]
        public void DeleteCncProgramWork(int CncProgramWorkId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "delete");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.DeleteCncProgramWork(CncProgramWorkId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteToolMachineParameter --機台所需刀具參數 刪除
        [HttpPost]
        public void DeleteToolMachineParameter(int ToolMachineId = -1)
        {
            try
            {
                WebApiLoginCheck("CncProgramLog", "delete");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.DeleteToolMachineParameter(ToolMachineId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//DeleteMillMachineTool --銑床刀具刪除
        [HttpPost]
        public void DeleteMillMachineTool(int MillMachineToolId = -1)
        {
            try
            {
                WebApiLoginCheck("MillMachineTool", "add");

                #region //Request
                masterCamDA = new MasterCamDA();
                dataRequest = masterCamDA.DeleteMillMachineTool(MillMachineToolId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Upload
        #endregion

        #region API
        #region //AddCncResponsestFile --新增編程回傳檔案--暫時用不到了
        //[HttpPost]
        //[Route("api/MES/CncResponsestFile")]
        //public void AddCncResponsestFile(string Company, string SecretKey, string FileName, string FileContent, string FileExtension, string FileSize, string ClientIP, string Source)
        //{
        //    try
        //    {
        //        #region //Api金鑰驗證
        //        ApiKeyVerify(Convert.ToInt32(Company), SecretKey, "AddCncResponsestFile");
        //        #endregion
        //        #region //Request
        //        dataRequest = masterCamDA.AddCncResponsestFile(Convert.ToInt32(Company), FileName, FileContent, FileExtension, FileSize, ClientIP, Source);
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
        #endregion

        #region//實體路徑轉Base64
        public void RealPathToBase(string result)
        {
            try
            {
                masterCamDA = new MasterCamDA();

                string ResponsestValues = "", CncParamNo = "", Values = "";
                int CncProgLogId = -1, FileId = -1;
                JObject ResultJson = JObject.Parse(result);
                CncProgLogId = int.Parse(ResultJson["NC_PROCESS_LOG_ID"].ToString());
                if (ResultJson["toolpath_url"] != null)
                {
                    for (int i = 0; i < ResultJson["toolpath_url"].Count(); i++)
                    {
                        ResponsestValues = ResultJson["toolpath_url"][i].ToString();
                        string[] NcData = ResponsestValues.Split(':');                       
                        FileId = ToBase(NcData[0].ToString());
                        CncParamNo = "toolpath_url";
                        dataRequest = masterCamDA.AddCncProgramResponsest(CncProgLogId, CncParamNo, ResponsestValues, FileId); //新增回傳資料
                    }
                }
                if (ResultJson["dxf"] != null)
                {
                    for (int i = 0; i < ResultJson["dxf"].Count(); i++)
                    {
                        ResponsestValues = ResultJson["dxf"][i].ToString();
                        FileId = ToBase(ResponsestValues);
                        CncParamNo = "dxf";
                        dataRequest = masterCamDA.AddCncProgramResponsest(CncProgLogId, CncParamNo, ResponsestValues, FileId); //新增回傳資料
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
        }
        #endregion

        #region //ToBase
        public int ToBase(string data) {
            int FileId = -1;
            try
            {
                masterCamDA = new MasterCamDA();

                string realPath = @"C:\Web\MES\Uploads\NC_PROGRAM_OUTPUT_PATH\output\" + data;                
                if (System.IO.File.Exists(realPath))
                {                    
                    var fileBinary = System.IO.File.ReadAllBytes(realPath);                   
                    string FileName = Path.GetFileNameWithoutExtension(realPath);
                    string FileExtension = Path.GetExtension(realPath);
                    FileInfo fileInfo = new FileInfo(realPath);
                    int FileSize = Convert.ToInt32(fileInfo.Length);
                    string Source = "/MasterCam/ViewCncProgramWorkLog";
                    string ClientIP = "::1";

                    #region //寫入資料庫
                    dataRequest = masterCamDA.AddCncResponsestFile(fileBinary, FileName, FileExtension, FileSize, ClientIP, Source);
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    FileId = int.Parse(jsonResponse["result"][0]["FileId"].ToString());
                    #endregion
                }
                else
                {
                    logger.Error("檔案不存在");
                }                
            }
            catch (Exception e) {
                logger.Error(e.ToString());
            }
            return FileId;
        }
        #endregion

        #region//Base64轉實體路徑
        public string BaseToRealPath(int FileId)
        {
            string result = "", error = "";
            try {
                masterCamDA = new MasterCamDA();

                string programPath = "C:\\Python\\Python36\\Data\\ZY_service_for_newdtm\\Bin2file.py";
                string FileJson = @"{'FileId':" + FileId + "}";
                FileJson = FileJson.Replace("'", "\"");

                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    Arguments = string.Format("\"{0}\" \"{1}\"", programPath, FileJson.Replace("\"", "\"\"")),
                    CreateNoWindow = true,
                    FileName = @"C:\Python\Python36\python.exe",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                };
                using (Process process = Process.Start(processStartInfo))
                {                    
                    using (StreamReader readerOutput = process.StandardOutput)
                    {
                        result = readerOutput.ReadToEnd().Replace("\r", "").Replace("\n", "");
                    }
                    if (result.Equals(""))
                    {
                        using (StreamReader readerError = process.StandardError)
                        {
                            error = readerError.ReadToEnd();
                            logger.Error(error);
                        }
                    }
                    if (error.Length > 0) result += error;
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
            return result;
        }
        #endregion

        #region//CopyNasFileToMes 將共夾檔案複製到IIS上
        public string CopyNasFileToMes(string NasFilePath)
        {
            string result = "", error = "";
            try
            {             
                if (System.IO.File.Exists(NasFilePath))
                {
                    string FileName = Path.GetFileNameWithoutExtension(NasFilePath);
                    string sourceFilePath = NasFilePath;
                    string destinationFolderPath = @"C:/Web/MES/Uploads/NC_PROGRAM_OUTPUT_PATH/output/";
                    string destinationFilePath = Path.Combine(destinationFolderPath, FileName);                    
                    System.IO.File.Copy(sourceFilePath, destinationFilePath,true);
                    result = destinationFilePath;
                }
                else {
                    throw new SystemException(NasFilePath+"【此檔案不存在】");
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
            return result;
        }
        #endregion

        #region//NC_Python 編程Python
        public string NC_Python(string data, string programPath)
        {
            string result = "", error = "";
            try {
                masterCamDA = new MasterCamDA();

                string ProgramPath = @programPath;
                data = data.Replace("'", @"""");
                data = data.Replace(@"""[{", @"[{");
                data = data.Replace(@"}]""", @"}]");

                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    Arguments = string.Format("\"{0}\" \"{1}\"", programPath, data.Replace("\"", "\"\"")),
                    CreateNoWindow = true,
                    FileName = @"C:\Python\Python36\python.exe",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false
                };
                using (Process process = Process.Start(processStartInfo))
                {                    
                    using (StreamReader readerOutput = process.StandardOutput)
                    {
                        result = readerOutput.ReadToEnd().Replace("\r", "").Replace("\n", "");
                        if (result.IndexOf("toolpath_url") != -1)
                        {
                            RealPathToBase(result);
                            logger.Error("status:success:" + result + ":" + data + ":" + programPath);
                        }
                        logger.Error("status:error:" + result + ":" + data + ":" + programPath);
                    }
                    if (result.Equals(""))
                    {
                        using (StreamReader readerError = process.StandardError)
                        {
                            error = readerError.ReadToEnd();
                            logger.Error("status:error:" + error +":"+ data+":"+ programPath);
                        }
                    }
                    if (error.Length > 0) result += error;                    
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
            return result;
        }
        #endregion

        #region//GetMoDwgInfo 取得設計圖API
        [HttpPost]
        [Route("api/ERP/GetMoDwgInfo")]
        public void GetMoDwgInfo(string CompanyNo = "", string SecretKey = "", int MoId = -1)
        {
            JObject jsonResponseNew = new JObject();
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetMoDwgInfo");
                #endregion

                #region //Request
                dataRequest = masterCamDA.GetMoDwgInfo(CompanyNo, MoId);
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
                    if (JObject.Parse(dataRequest)!= null)
                    {
                        jsonResponseNew = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "OK",
                            DwgResult = JObject.Parse(dataRequest)["data1"],
                            JigResult = JObject.Parse(dataRequest)["data2"],
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
                jsonResponseNew = JObject.FromObject(new
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

        #region//GetMillMachineTool 取得銑床機台刀具
        [HttpPost]
        [Route("api/ERP/GetMillMachineTool")]
        public void GetMillMachineToolApi(string CompanyNo = "", string SecretKey = "", string MachineNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetMillMachineTool");
                #endregion

                #region //Request
                dataRequest = masterCamDA.GetMillMachineTool("",-1,-1,-1, MachineNo, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetMoProcessApi 取得製令製程清單
        [HttpPost]
        [Route("api/ERP/GetMoProcessApi")]
        public void GetMoProcessApi(string CompanyNo = "", string SecretKey = "", string JigNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetMoProcess");
                #endregion

                #region //Request
                dataRequest = masterCamDA.GetJigMoProcess(JigNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddMillPreEditLog 新增銑床預編Log & 預編Request資訊 
        [HttpPost]
        [Route("api/ERP/AddMillPreEditLog")]
        public void AddMillPreEditLog(string CompanyNo = "", string SecretKey = "", int MoProcessId = -1,string JigNo="", string DwgData = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "AddMillPreEditLog");
                #endregion

                #region //Request
                masterCamDA = new MasterCamDA();
                GetCadFloderPath(2);
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = masterCamDA.AddMillPreEditLog(MoProcessId, JigNo, ServerPath, DwgData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//GetPreData 取得預編檔案路徑
        [HttpPost]
        [Route("api/ERP/GetPreData")]
        public void GetPreData(string CompanyNo = "", string SecretKey = "", int MillPreLogId = -1,string JigNo="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "GetPreData");
                #endregion

                #region //Request
                dataRequest = masterCamDA.GetPreData(MillPreLogId,JigNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddMillPreEditResponsest 新增銑床預編回傳紀錄
        [HttpPost]
        [Route("api/ERP/AddMillPreEditResponsest")]
        public void AddMillPreEditResponsest(string CompanyNo = "", string SecretKey = "", int MillPreLogId = -1,string MillPrePrtData="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "AddMillPreEditResponsest");
                #endregion

                #region //Request
                masterCamDA = new MasterCamDA();
                GetCadFloderPath(2);
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = masterCamDA.AddMillPreEditResponsest(MillPreLogId, CompanyNo, ServerPath, MillPrePrtData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddMillProgramLog 新增銑床編程需求紀錄
        [HttpPost]
        [Route("api/ERP/AddMillProgramLog")]
        public void AddMillProgramLog(string CompanyNo = "", string SecretKey = "", int MillPreLogId = -1
            , string MachineNo = "", string MillToolData="",string MillPrePrtData="",string JigNo="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "AddMillProgramLog");
                #endregion

                #region //Request
                masterCamDA = new MasterCamDA();
                GetCadFloderPath(2);
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = masterCamDA.AddMillProgramLog(CompanyNo,SecretKey, MillPreLogId, MachineNo, MillToolData,MillPrePrtData,ServerPath,JigNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMillProgramResponsest 新增銑床編程回傳紀錄
        [HttpPost]
        [Route("api/ERP/AddMillProgramResponsest")]
        public void AddMillProgramResponsest(string CompanyNo = "", string SecretKey = "", int MillProgLogId = -1, string MillPrtData = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "AddMillProgramResponsest");
                masterCamDA = new MasterCamDA();
                #endregion

                #region //Request
                GetCadFloderPath(2);
                string ServerPath = Server.MapPath(CadFolderPath);
                dataRequest = masterCamDA.AddMillProgramResponsest(MillProgLogId, CompanyNo, ServerPath, MillPrtData);
                #endregion               

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//AddMillProgramWork 新增銑床編程報工紀錄
        [HttpPost]
        [Route("api/ERP/AddMillProgramWork")]
        public void AddMillProgramWork(string CompanyNo = "", string SecretKey = "", int MillProgLogId = -1, int MoProcessId = -1
            ,string MachineNo = "",string StartWorkDate="",string EndWorkDate="")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "AddMillProgramWork");
                #endregion

                #region //Request
                masterCamDA = new MasterCamDA();               
                dataRequest = masterCamDA.AddMillProgramWork(MillProgLogId, MoProcessId, MachineNo, StartWorkDate, EndWorkDate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region//客戶端 檔案上傳測試
        [HttpPost]
        [Route("api/ERP/clientTest")]
        public async Task<string> clientTest(string STEP_DATA="")
        {
            var clientUrl = "http://localhost:56619/api/ERP/getTest";
            //var clientUrl = "http://192.168.145.55:8080/api/YcmNCData";
            HttpClient httpClient = new HttpClient();
            HttpClient client = new HttpClient();
            byte[] FileBytes = System.IO.File.ReadAllBytes(STEP_DATA);
            MultipartFormDataContent formData = new MultipartFormDataContent
            {                               
                { new StreamContent(new MemoryStream(FileBytes)), "PostNCData",Path.GetFileName(STEP_DATA)}
            };
            var result = await client.PostAsync(clientUrl, formData);
            if (result.IsSuccessStatusCode)
            {
                var response = new HttpResponseMessage(HttpStatusCode.OK);
            }
            else
            {
                var response = new HttpResponseMessage(HttpStatusCode.NotFound);
            }
            string res = "";
            return res;
        }
        #endregion

        #region//客戶端 檔案上傳測試
        [HttpPost]
        [Route("api/ERP/getTest")]
        public string FCFileUpload()
        {
            GetCadFloderPath(2);
            string ServerPath = Server.MapPath(CadFolderPath);
            foreach (string uploadFile in Request.Files)
            {
                string originalFileName = Path.GetFileNameWithoutExtension(Request.Files[uploadFile].FileName);
                string originalFileExt = Path.GetExtension(Request.Files[uploadFile].FileName);

                Request.Files[uploadFile].SaveAs(Path.Combine(ServerPath, originalFileName + originalFileExt));
            }
            return "A";
        }
        #endregion
        #endregion
    }
}