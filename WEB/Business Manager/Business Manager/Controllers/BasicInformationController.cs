using BASDA;
using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class BasicInformationController : WebController
    {
        string UploadFolderPath = @"";

        #region //取得共夾檔案上傳路徑
        public void GetUploadPath()
        {
            int CompanyId = Convert.ToInt32(System.Web.HttpContext.Current.Session["UserCompany"]);

            if (System.Web.HttpContext.Current.Session["CompanySwitch"] != null)
            {
                if (System.Web.HttpContext.Current.Session["CompanySwitch"].ToString() == "manual")
                {
                    CompanyId = Convert.ToInt32(System.Web.HttpContext.Current.Session["CompanyId"]);
                }
            }

            switch (CompanyId)
            {
                case 2:
                    UploadFolderPath = @"~/Uploads";
                    break;
                case 4:
                    UploadFolderPath = @"~/Uploads";
                    break;
                case 3:
                    UploadFolderPath = @"~/Uploads";
                    break;
                default:
                    throw new SystemException("此公司別尚未維護上傳路徑!!");
            }
        }
        #endregion

        #region //View
        public ActionResult UITemplate()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult CompanyManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult DepartmentManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UserManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SystemSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ModuleSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult FunctionSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult StatusSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TypeSetting()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ShiftManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AssetsManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult UploadWhitelistManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetStatus 取得狀態資料
        [HttpPost]
        public void GetStatus(string StatusSchema = "", string StatusNo = "", string StatusName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetStatus(StatusSchema, StatusNo, StatusName
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

        #region //GetStatusSchema 取得狀態綱要
        [HttpPost]
        public void GetStatusSchema(string StatusSchema = "")
        {
            try
            {
                WebApiLoginCheck("StatusSetting", "read");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetStatusSchema(StatusSchema);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetStatusCompany 取得狀態公司對應資料
        [HttpPost]
        public void GetStatusCompany(int StatusId = -1)
        {
            try
            {
                WebApiLoginCheck("StatusSetting", "status-company");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetStatusCompany(StatusId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetType 取得類別資料
        [HttpPost]
        public void GetType(string TypeSchema = "", string TypeNo = "", string TypeName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetType(TypeSchema, TypeNo, TypeName
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

        #region //GetTypeSchema 取得類別綱要
        [HttpPost]
        public void GetTypeSchema(string TypeSchema = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TypeSetting", "read");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetTypeSchema(TypeSchema
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

        #region //GetTypeCompany 取得類別公司對應資料
        [HttpPost]
        public void GetTypeCompany(int TypeId = -1)
        {
            try
            {
                WebApiLoginCheck("TypeSetting", "type-company");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetTypeCompany(TypeId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCompany 取得公司資料
        [HttpPost]
        public void GetCompany(int CompanyId = -1, string CompanyNo = "", string CompanyName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CompanyManagement", "read,constrained-data");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetCompany(CompanyId, CompanyNo, CompanyName, Status
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

        #region //GetDepartment 取得部門資料
        [HttpPost]
        public void GetDepartment(string CompanyNo = "", int DepartmentId = -1, int CompanyId = -1, string DepartmentNo = "", string DepartmentName = "", string Status = ""
            , string SearchKey = "", string SearchKeys = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "read,constrained-data");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetDepartment(CompanyNo, DepartmentId, CompanyId, DepartmentNo, DepartmentName, Status
                    , SearchKey, SearchKeys
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

        #region //GetDepartmentShift 取得部門班次資料
        [HttpPost]
        public void GetDepartmentShift(int DepartmentId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "department-shift");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetDepartmentShift(DepartmentId
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

        #region //GetDepartmentRate 取得部門費率資料 -- Chia Yuan --230628
        [HttpPost]
        public void GetDepartmentRate(int DepartmentId = -1, int AuthorId = -1, string Status = ""
            , string StartEnableDate = "", string EndEnableDate = ""
            , string StartDisabledDate = "", string EndDisabledDate = ""
            , string StartCreateDate = "", string EndCreateDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "department-rate");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetDepartmentRate(DepartmentId, AuthorId, Status
                    , StartEnableDate, EndEnableDate
                    , StartDisabledDate, EndDisabledDate
                    , StartCreateDate, EndCreateDate
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

        #region //GetUser 取得使用者資料
        [HttpPost]
        public void GetUser(int UserId = -1, int DepartmentId = -1, int CompanyId = -1, string Departments = ""
            , string UserNo = "", string UserName = "", string Gender = "", string Status = "", string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserManagement", "read,constrained-data");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetUser(UserId, DepartmentId, CompanyId, Departments
                    , UserNo, UserName, Gender, Status, SearchKey
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

        #region //GetUserByAuthority 取得使用者資料(權限)
        [HttpPost]
        public void GetUserByAuthority(int UserId = -1, int DepartmentId = -1, int CompanyId = -1, string Departments = ""
            , string UserNo = "", string UserName = "", string Gender = "", string Status = "", string SearchKey = ""
            , string FunctionCode = "", string DetailCode = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UserManagement", "read,constrained-data");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetUserByAuthority(UserId, DepartmentId, CompanyId, Departments
                    , UserNo, UserName, Gender, Status, SearchKey
                    , FunctionCode, DetailCode
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

        #region //GetSystem 取得系統別資料
        [HttpPost]
        public void GetSystem(int SystemId = -1, string SystemCode = "", string SystemName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetSystem(SystemId, SystemCode, SystemName, Status
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

        #region //GetModule 取得模組別資料
        [HttpPost]
        public void GetModule(int ModuleId = -1, int SystemId = -1, string ModuleCode = "", string ModuleName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetModule(ModuleId, SystemId, ModuleCode, ModuleName, Status
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

        #region //GetFunction 取得功能別資料
        [HttpPost]
        public void GetFunction(int FunctionId = -1, int ModuleId = -1, int SystemId = -1, string FunctionCode = "", string FunctionName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetFunction(FunctionId, ModuleId, SystemId, FunctionCode, FunctionName, Status
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

        #region //GetFunctionDetail 取得詳細功能別資料
        [HttpPost]
        public void GetFunctionDetail(int DetailId = -1, int FunctionId = -1, string DetailCode = "", string DetailName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "read");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetFunctionDetail(DetailId, FunctionId, DetailCode, DetailName, Status
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

        #region //GetFunctionDetailCode 取得功能詳細代碼
        [HttpPost]
        public void GetFunctionDetailCode()
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "read");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetFunctionDetailCode();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSystemMenu 取得系統選單資料
        [HttpPost]
        public void GetSystemMenu()
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetSystemMenu();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetModuleMenu 取得模組選單資料
        [HttpPost]
        public void GetModuleMenu(string SystemCode = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetModuleMenu(SystemCode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetShift 取得班次資料
        [HttpPost]
        public void GetShift(int ShiftId = -1, string ShiftName = "", string WorkBeginTime = "", string WorkEndTime = "", string WorkHours = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ShiftManagement", "read");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetShift(ShiftId, ShiftName, WorkBeginTime, WorkEndTime, WorkHours
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

        #region //GetDepartmentRateDetail 取部門得費用率詳細資料 -- Chia Yuan --230627
        [HttpPost]
        public void GetDepartmentRateDetail(int DepartmentRateId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ShiftManagement", "read");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetDepartmentRateDetail(DepartmentRateId
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

        #region //GetErpData 取得Erp資料
        [HttpPost]
        public void GetErpData(string UserNo = "", string ColumnName = "", string Condition = ""
            , string OrderBy = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetErpData(UserNo, ColumnName, Condition, OrderBy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetErpSysData 取得ErpSys資料
        [HttpPost]
        public void GetErpSysData(string ColumnName = "", string Condition = ""
            , string OrderBy = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetErpSysData(ColumnName, Condition, OrderBy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetAssetsData 取得ERP資產資料
        [HttpPost]
        public void GetAssetsData(string UserNo = "", string MB001 = "", string MB002 = "", string MB006 = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetAssetsData(UserNo, MB001, MB002, MB006
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

        #region //GetUploadWhitelist 取得上傳路徑白名單
        [HttpPost]
        public void GetUploadWhitelist(int ListId = -1, string ListNo = "", string FolerPath = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetUploadWhitelist(ListId, ListNo, FolerPath);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCheckUploadWhitelist 確認路徑是否為合法路徑
        [HttpPost]
        public void GetCheckRdWhitelist(int ListId = -1, string ListNo = "", string FolderPath = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetCheckUploadWhitelist(ListId, ListNo, FolderPath);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetFolder 取得資料夾項目
        [HttpPost]
        public void GetFolder(string FolderPath = "", int ListId = -1, string ListNo = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //確認此路徑為合法路徑
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetCheckUploadWhitelist(ListId, ListNo, FolderPath);

                var dataRequestJson = JObject.Parse(dataRequest);
                foreach (var item in dataRequestJson)
                {
                    if (item.Key == "status")
                    {
                        if (item.Value.ToString() == "success")
                        {
                            break;
                        }
                    }
                    else if (item.Key == "msg")
                    {
                        throw new SystemException(item.Value.ToString());
                    }
                }
                #endregion

                string[] Directories = Directory.GetDirectories(FolderPath);

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
        public void GetFiles(string FolderPath = "", int ListId = -1, string ListNo = "")
        {
            try
            {
                WebApiLoginCheck();

                string[] Files = Directory.GetFiles(FolderPath);
                #region //確認白名單資料是否正確
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetCheckUploadWhitelist(ListId, ListNo, FolderPath);

                jsonResponse = BaseHelper.DAResponse(dataRequest);

                string WhitelistPath = "";
                if (jsonResponse["status"].ToString() == "success")
                {
                    WhitelistPath = jsonResponse["result"][0]["FolderPath"].ToString();
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
                #endregion

                List<FilePathInfo> filePaths = new List<FilePathInfo>();
                foreach (var files in Files)
                {
                    var file = new FileInfo(files);
                    string FilePath = file.ToString();

                    #region //處理URL特殊符號
                    if (FilePath.IndexOf("+") != -1) FilePath = FilePath.Replace("+", "%2B");
                    if (FilePath.IndexOf("/") != -1) FilePath = FilePath.Replace("/", "%2F");
                    if (FilePath.IndexOf("?") != -1) FilePath = FilePath.Replace("?", "%3F");
                    if (FilePath.IndexOf("#") != -1) FilePath = FilePath.Replace("#", "%23");
                    if (FilePath.IndexOf("&") != -1) FilePath = FilePath.Replace("&", "%26");
                    if (FilePath.IndexOf("=") != -1) FilePath = FilePath.Replace("=", "%3D");
                    #endregion

                    var fileInfo = new FilePathInfo
                    {
                        FileName = file.Name,
                        FilePath = FilePath,
                        FolderPath = FolderPath,
                        WhitelistPath = WhitelistPath
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

        #region //GetFileInfo 取得檔案相關資訊
        [HttpPost]
        public void GetFileInfo(string FilePath = "", string MESFolderPath = "", string DowloadFlag = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetFileInfo(FilePath, MESFolderPath, DowloadFlag);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUploadFile 取得檔案(絕對路徑方式)
        public virtual ActionResult GetUploadFile(string FilePath = "", string DownloadFlag = "")
        {
            try
            {
                if (BaseHelper.ClientLinkType() == "廠內連線")
                {
                    #region //Request
                    GetUploadPath();
                    string ServerPath = Server.MapPath(UploadFolderPath);
                    dataRequest = basicInformationDA.GetFileInfo(FilePath, ServerPath, DownloadFlag);
                    #endregion

                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
                    #endregion

                    if (jsonResponse["status"].ToString() == "success")
                    {
                        string fileName = jsonResponse["result"][0]["FileName"].ToString();
                        byte[] fileContent = (byte[])jsonResponse["result"][0]["FileByte"];
                        string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();

                        return File(fileContent, FileHelper.GetMime(fileExtension), fileName + fileExtension);
                    }
                    else
                    {
                        throw new SystemException(jsonResponse["msg"].ToString());
                    }
                }
                else
                {
                    throw new SystemException("此IP非廠內路徑，不可以下載");
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
                Response.ContentType = "application/json";
                Response.Write(jsonResponse.ToString(Newtonsoft.Json.Formatting.None));
                return new EmptyResult();
            }
        }
        #endregion

        #region //GetUploadWhitelistManagement 取得白名單上傳路徑資料
        [HttpPost]
        public void GetUploadWhitelistManagement(int ListId = -1, string ListNo = "", int DepartmentId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("UploadWhitelistManagement", "read,constrained-data");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetUploadWhitelistManagement(ListId, ListNo, DepartmentId
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
        #region //AddCompany 公司資料新增
        [HttpPost]
        public void AddCompany(string CompanyNo = "", string CompanyName = "", string Telephone = "", string Fax = "", string Address = ""
            , string AddressEn = "", int LogoIcon = -1)
        {
            try
            {
                WebApiLoginCheck("CompanyManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddCompany(CompanyNo, CompanyName, Telephone, Fax, Address
                    , AddressEn, LogoIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDepartment 部門資料新增
        [HttpPost]
        public void AddDepartment(int CompanyId = -1, string DepartmentNo = "", string DepartmentName = "")
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddDepartment(CompanyId, DepartmentNo, DepartmentName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDepartmentShift 部門班次新增
        [HttpPost]
        public void AddDepartmentShift(int DepartmentId = -1, string ShiftId = "")
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddDepartmentShift(DepartmentId, ShiftId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddDepartmentRate 部門費用率新增 -- Chia Yuan --230627

        [HttpPost]
        public void AddDepartmentRate(int DepartmentId = -1, int AuthorId = -1, decimal ResourceRate = 0, decimal OverheadRate = 0)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddDepartmentRate(DepartmentId, AuthorId, ResourceRate, OverheadRate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //AddUser 使用者資料新增
        [HttpPost]
        public void AddUser(int DepartmentId = -1, string UserNo = "", string UserName = "", string Gender = "", string Email = "")
        {
            try
            {
                WebApiLoginCheck("UserManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddUser(DepartmentId, UserNo, UserName, Gender, Email);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddSystem 系統別資料新增
        [HttpPost]
        public void AddSystem(string SystemCode = "", string SystemName = "", string ThemeIcon = "")
        {
            try
            {
                WebApiLoginCheck("SystemSetting", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddSystem(SystemCode, SystemName, ThemeIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddModule 模組別資料新增
        [HttpPost]
        public void AddModule(int SystemId = -1, string ModuleCode = "", string ModuleName = "", string ThemeIcon = "")
        {
            try
            {
                WebApiLoginCheck("ModuleSetting", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddModule(SystemId, ModuleCode, ModuleName, ThemeIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddFunction 功能別資料新增
        [HttpPost]
        public void AddFunction(int ModuleId = -1, string FunctionCode = "", string FunctionName = "", string UrlTarget = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddFunction(ModuleId, FunctionCode, FunctionName, UrlTarget);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddFunctionDetail 詳細功能別資料新增
        [HttpPost]
        public void AddFunctionDetail(int FunctionId = -1, string DetailCode = "", string DetailName = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddFunctionDetail(FunctionId, DetailCode, DetailName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddShift 班次資料新增
        [HttpPost]
        public void AddShift(string ShiftName = "", string WorkBeginTime = "", string WorkEndTime = "", string WorkHours = "")
        {
            try
            {
                WebApiLoginCheck("ShiftManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddShift(ShiftName, WorkBeginTime, WorkEndTime, WorkHours);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddUploadWhitelistManagement 白名單上傳新增
        [HttpPost]
        public void AddUploadWhitelistManagement(string ListNo = "", int DepartmentId = -1, string FolderPath = "")
        {
            try
            {
                WebApiLoginCheck("UploadWhitelistManagement", "add");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.AddUploadWhitelistManagement(ListNo, DepartmentId, FolderPath);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateCompany 公司資料更新
        [HttpPost]
        public void UpdateCompany(int CompanyId = -1, string CompanyName = "", string Telephone = "", string Fax = "", string Address = ""
            , string AddressEn = "", int LogoIcon = -1)
        {
            try
            {
                WebApiLoginCheck("CompanyManagement", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateCompany(CompanyId, CompanyName, Telephone, Fax, Address
                    , AddressEn, LogoIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCompanyStatus 公司狀態更新
        [HttpPost]
        public void UpdateCompanyStatus(int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("CompanyManagement", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateCompanyStatus(CompanyId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDepartment 部門資料更新
        [HttpPost]
        public void UpdateDepartment(int DepartmentId = -1, int CompanyId = -1, string DepartmentName = "")
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateDepartment(DepartmentId, CompanyId, DepartmentName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDepartmentStatus 部門狀態更新
        [HttpPost]
        public void UpdateDepartmentStatus(int DepartmentId = -1)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateDepartmentStatus(DepartmentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateDepartmentRateStatus 部門費用率狀態更新 -- Chia Yuan --230627

        [HttpPost]
        public void UpdateDepartmentRateStatus(int DepartmentRateId = -1)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateDepartmentRateStatus(DepartmentRateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateDepartmentRate 部門費用率更新 -- Chia Yuan --230628

        [HttpPost]
        public void UpdateDepartmentRate(int DepartmentRateId = -1, int AuthorId = -1, decimal ResourceRate = 0, decimal OverheadRate = 0)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateDepartmentRate(DepartmentRateId, AuthorId, ResourceRate, OverheadRate);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion

        #region //UpdateUser 使用者資料更新
        [HttpPost]
        public void UpdateUser(int UserId = -1, int DepartmentId = -1, string UserName = "", string Gender = "", string Email = "")
        {
            try
            {
                WebApiLoginCheck("UserManagement", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateUser(UserId, DepartmentId, UserName, Gender, Email);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUserStatus 使用者狀態更新
        [HttpPost]
        public void UpdateUserStatus(int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("UserManagement", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateUserStatus(UserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSystem 系統別資料更新
        [HttpPost]
        public void UpdateSystem(int SystemId = -1, string SystemCode = "", string SystemName = "", string ThemeIcon = "")
        {
            try
            {
                WebApiLoginCheck("SystemSetting", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateSystem(SystemId, SystemCode, SystemName, ThemeIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSystemStatus 系統別狀態更新
        [HttpPost]
        public void UpdateSystemStatus(int SystemId = -1)
        {
            try
            {
                WebApiLoginCheck("SystemSetting", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateSystemStatus(SystemId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSystemSort 系統別順序調整
        [HttpPost]
        public void UpdateSystemSort(string SystemList = "")
        {
            try
            {
                WebApiLoginCheck("SystemSetting", "sort");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateSystemSort(SystemList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateModule 模組別資料更新
        [HttpPost]
        public void UpdateModule(int ModuleId = -1, int SystemId = -1, string ModuleCode = "", string ModuleName = "", string ThemeIcon = "")
        {
            try
            {
                WebApiLoginCheck("ModuleSetting", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateModule(ModuleId, SystemId, ModuleCode, ModuleName, ThemeIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateModuleStatus 模組別狀態更新
        [HttpPost]
        public void UpdateModuleStatus(int ModuleId = -1)
        {
            try
            {
                WebApiLoginCheck("ModuleSetting", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateModuleStatus(ModuleId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateModuleSort 模組別順序調整
        [HttpPost]
        public void UpdateModuleSort(int SystemId = -1, string ModuleList = "")
        {
            try
            {
                WebApiLoginCheck("ModuleSetting", "sort");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateModuleSort(SystemId, ModuleList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFunction 功能別資料更新
        [HttpPost]
        public void UpdateFunction(int FunctionId = -1, int ModuleId = -1, string FunctionCode = "", string FunctionName = "", string UrlTarget = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateFunction(FunctionId, ModuleId, FunctionCode, FunctionName, UrlTarget);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFunctionStatus 功能別狀態更新
        [HttpPost]
        public void UpdateFunctionStatus(int FunctionId = -1)
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateFunctionStatus(FunctionId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFunctionSort 功能別順序調整
        [HttpPost]
        public void UpdateFunctionSort(int ModuleId = -1, string FunctionList = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "sort");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateFunctionSort(ModuleId, FunctionList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFunctionDetail 詳細功能別資料更新
        [HttpPost]
        public void UpdateFunctionDetail(int DetailId = -1, int FunctionId = -1, string DetailCode = "", string DetailName = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateFunctionDetail(DetailId, FunctionId, DetailCode, DetailName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFunctionDetailStatus 詳細功能別狀態更新
        [HttpPost]
        public void UpdateFunctionDetailStatus(int DetailId = -1)
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "status-switch");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateFunctionDetailStatus(DetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFunctionDetailSort 詳細功能別順序調整
        [HttpPost]
        public void UpdateFunctionDetailSort(int FunctionId = -1, string FunctionDetailList = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "update, sort");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateFunctionDetailSort(FunctionId, FunctionDetailList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateStatus 狀態資料更新
        [HttpPost]
        public void UpdateStatus(string StatusSchema = "", string StatusNo = "", string StatusName = "", string StatusDesc = "")
        {
            try
            {
                WebApiLoginCheck("StatusSetting", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateStatus(StatusSchema, StatusNo, StatusName, StatusDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateStatusCompany 狀態公司對應資料更新
        [HttpPost]
        public void UpdateStatusCompany(int StatusId = -1, int CompanyId = -1, string StatusName = "", string StatusDesc = "")
        {
            try
            {
                WebApiLoginCheck("StatusSetting", "status-company");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateStatusCompany(StatusId, CompanyId, StatusName, StatusDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateType 類別資料更新
        [HttpPost]
        public void UpdateType(string TypeSchema = "", string TypeNo = "", string TypeName = "", string TypeDesc = "")
        {
            try
            {
                WebApiLoginCheck("TypeSetting", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateType(TypeSchema, TypeNo, TypeName, TypeDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTypeCompany 類別公司對應資料更新
        [HttpPost]
        public void UpdateTypeCompany(int TypeId = -1, int CompanyId = -1, string TypeName = "", string TypeDesc = "")
        {
            try
            {
                WebApiLoginCheck("TypeSetting", "type-company");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateTypeCompany(TypeId, CompanyId, TypeName, TypeDesc);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateShift 班次資料更新
        [HttpPost]
        public void UpdateShift(int ShiftId = -1, string ShiftName = "", string WorkBeginTime = "", string WorkEndTime = "", string WorkHours = "")
        {
            try
            {
                WebApiLoginCheck("ShiftManagement", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateShift(ShiftId, ShiftName, WorkBeginTime, WorkEndTime, WorkHours);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateSendResetMail 寄送使用者忘記密碼信件
        [HttpPost]
        public void UpdateSendResetMail(string UserNo)
        {
            try
            {
                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateSendResetMail(UserNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUploadWhitelistManagement 白名單上傳更新
        [HttpPost]
        public void UpdateUploadWhitelistManagement(int ListId = -1, string ListNo = "", int DepartmentId = -1, string FolderPath = "")
        {
            try
            {
                WebApiLoginCheck("UploadWhitelistManagement", "update");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.UpdateUploadWhitelistManagement(ListId, ListNo, DepartmentId, FolderPath);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteStatus 狀態資料刪除
        [HttpPost]
        public void DeleteStatus(string StatusSchema = "", string StatusNo = "")
        {
            try
            {
                WebApiLoginCheck("StatusSetting", "delete");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.DeleteStatus(StatusSchema, StatusNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteType 類別資料刪除
        [HttpPost]
        public void DeleteType(string TypeSchema = "", string TypeNo = "")
        {
            try
            {
                WebApiLoginCheck("TypeSetting", "delete");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.DeleteType(TypeSchema, TypeNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteShift 班次資料刪除
        [HttpPost]
        public void DeleteShift(int ShiftId = -1)
        {
            try
            {
                WebApiLoginCheck("ShiftManagement", "delete");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.DeleteShift(ShiftId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteShift 部門費用率詳細資料刪除 --Chia Yuan --230627
        [HttpPost]
        public void DeleteDepartmentRateDetail(int DepartmentRateId = -1)
        {
            try
            {
                WebApiLoginCheck("ShiftManagement", "delete");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.DeleteDepartmentRateDetail(DepartmentRateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteDepartmentShift 部門班次資料刪除
        [HttpPost]
        public void DeleteDepartmentShift(int DepartmentId = -1, int ShiftId = -1)
        {
            try
            {
                WebApiLoginCheck("DepartmentManagement", "department-shift");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.DeleteDepartmentShift(DepartmentId, ShiftId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteUploadWhitelistManagement 白名單上傳刪除
        [HttpPost]
        public void DeleteUploadWhitelistManagement(int ListId = -1)
        {
            try
            {
                WebApiLoginCheck("UploadWhitelistManagement", "delete");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.DeleteUploadWhitelistManagement(ListId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UserExcelDownload
        public void UserExcelDownload(string Departments = "", int CompanyId = -1, string UserNo = "", string UserName = ""
            , string Gender = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("UserManagement", "excel");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetUser(-1, -1, CompanyId, Departments
                    , UserNo, UserName, Gender, Status, ""
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
                    headerStyle.Fill.BackgroundColor = XLColor.AppleGreen;
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
                    string excelFileName = "使用者管理Excel檔";
                    string excelsheetName = "使用者詳細資料";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] headerValues = new string[] { "公司", "部門", "職務", "職務類別", "使用者編號", "使用者名稱", "性別", "Email" };
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
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.CompanyNo.ToString() + "-" + item.CompanyName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.DepartmentNo.ToString() + "-" + item.DepartmentName.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.Job.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.JobType.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.UserNo.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.UserName.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = "(" + item.Gender.ToString() + ")" + item.UserName.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).RichText.Substring(0, 3).SetFontColor(XLColor.Red).SetBold().SetStrikethrough();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.Gender.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.Email.ToString();

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = "2022/05/18 13:27:12";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Style = dateStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = "1234560";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Style = numberStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = "123000";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Style = currencyStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).FormulaA1 = "=IF(" + BaseHelper.MergeNumberToChar(7, rowIndex) + "=\"F\",1,0)";
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
                        worksheet.Protect("123");

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

        #region //AssetsDataExcelDownload
        public void AssetsDataExcelDownload(string UserNo = "", string MB001 = "", string MB002 = "", string MB006 = "")
        {
            try
            {
                WebApiLoginCheck("AssetsManagement", "excel");

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetAssetsData(UserNo, MB001, MB002, MB006
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
                    string excelFileName = "資產清單Excel檔";
                    string excelsheetName = "資產清單詳細資料";

                    dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequest)["data"].ToString());
                    string[] headerValues = new string[] { "部門代號", "部門名稱", "資產類別", "資產編號", "資產名稱", "資產規格", "保管人姓名", "數量", "取得日期", "取得成本" };
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
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = item.部門代號.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = item.部門名稱.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = item.資產類別.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = item.資產編號.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = item.資產名稱.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = item.資產規格.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = item.保管人工號.ToString() + item.保管人姓名.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = item.數量.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = item.取得日期.ToString();
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = item.取得成本.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = "(" + item.Gender.ToString() + ")" + item.UserName.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).RichText.Substring(0, 3).SetFontColor(XLColor.Red).SetBold().SetStrikethrough();

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = "2022/05/18 13:27:12";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Style = dateStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = "1234560";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Style = numberStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Value = "123000";
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).Style = currencyStyle;

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).FormulaA1 = "=IF(" + BaseHelper.MergeNumberToChar(7, rowIndex) + "=\"F\",1,0)";
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

        #region //FOR EIP API

        #region //GetTypeEIP 取得類別資料
        [HttpPost]
        [Route("api/BasicInformation/GetTypeEIP")]

        public void GetTypeEIP(string TypeSchema = "", string TypeNo = "", string TypeName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1, int[] CustomerIds = null)
        {
            try
            {
               // WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetTypeEIP(TypeSchema, TypeNo, TypeName
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

        #region //GetStatusEIP 取得狀態資料
        [HttpPost]
        [Route("api/BasicInformation/GetStatusEIP")]
        public void GetStatusEIP(string StatusSchema = "", string StatusNo = "", string StatusName = ""
            , int[] CustomerIds = null
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck();

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetStatusEIP(StatusSchema, StatusNo, StatusName
                    , CustomerIds
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
    }
}