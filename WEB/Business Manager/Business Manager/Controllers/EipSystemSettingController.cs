using EIPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class EipSystemSettingController : WebController
    {
        public EipSystemSettingDA eipSystemSettingDA = new EipSystemSettingDA();

        #region //View
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

        public ActionResult AuthorityManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult CorpCustManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult MemberManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetFunctionDetailCode 取得功能詳細代碼
        [HttpPost]
        public void GetFunctionDetailCode()
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetFunctionDetailCode();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void GetFunctionDetail(int FnDetailId = -1, int FunctionId = -1, string FnDetailCode = "", string FnDetailName = "", string Status = "", int[] CustomerIds = null
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetFunctionDetail(FnDetailId, FunctionId, FnDetailCode, FnDetailName, Status
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

        #region //GetRole 取得角色資料
        [HttpPost]
        public void GetRole(int RoleId = -1, string RoleName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetRole(RoleId, RoleName, Status
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

        #region //GetAuthorityMember 取得權限會員資料
        [HttpPost]
        public void GetAuthorityMember(string Roles = "", string MemberEmail = "", string MemberName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetAuthorityMember(Roles, MemberEmail, MemberName, Status
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

        #region //GetCsCustomer 會員客戶對照表
        [HttpPost]
        public void GetCsCustomer(int CsCustId = -1, string CustomerName = "", string CustomerEnglishName = "", string SearchKey = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetCsCustomer(CsCustId, CustomerName, CustomerEnglishName, SearchKey, Status
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

        #region //GetCustomer 取得客戶資料表
        [HttpPost]
        public void GetCustomer(int CustomerId = -1, int CompanyId = -1, string CustomerName = "", string CustomerEnglishName = "", string SearchKey = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetCustomer(CustomerId, CompanyId, CustomerName, CustomerEnglishName, SearchKey, Status
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

        #region //CsCustomerDetail 客戶明細表
        [HttpPost]
        public void GetCsCustomerDetail(int CsCustId = -1, int CustomerId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetCsCustomerDetail(CsCustId, CustomerId, Status
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

        #region //GetMember -- 取得會員資料
        [HttpPost]
        public void GetMember(int MemberId = -1, int CsCustId = -1, string MemberEmail = "", string MemberName = "", string Status = ""
            , string Gender = "", string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetMember(MemberId, CsCustId, MemberEmail, MemberName, Status
                    , Gender, SearchKey
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

        #region //GetRoleAuthority 取得角色權限資料
        [HttpPost]
        public void GetRoleAuthority(int RoleId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetRoleAuthority(RoleId, SearchKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRoleDetailAuthority 取得角色詳細權限資料
        [HttpPost]
        public void GetRoleDetailAuthority(int ModuleId = -1, int RoleId = -1, string SearchKey = "", bool Grant = false)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetRoleDetailAuthority(ModuleId, RoleId, SearchKey, Grant);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRoleUser 取得角色會員資料
        [HttpPost]
        public void GetRoleUser(int RoleId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetRoleUser(RoleId, SearchKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserAuthority 取得會員權限資料
        [HttpPost]
        public void GetUserAuthority(int MemberId = -1, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetUserAuthority(MemberId, SearchKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserDetailAuthority 取得使用者詳細權限資料
        [HttpPost]
        public void GetUserDetailAuthority(int ModuleId = -1, int MemberId = -1, int RoleId = -1, string SearchKey = "", bool Grant = false)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetUserDetailAuthority(ModuleId, MemberId, RoleId, SearchKey, Grant);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserRole 取得使用者角色資料
        [HttpPost]
        public void GetUserRole(string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetUserRole(SearchKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMemberCustCorp 取得會員客戶公司資料
        [HttpPost]
        public void GetMemberCustCorp(string CompanyIds = "", string CsCustIds = "", string MemberEmail = "", string MemberName = "", string Status = "", string CsCustStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetMemberCustCorp(CompanyIds, CsCustIds, MemberEmail, MemberName, Status, CsCustStatus
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

        #region //GetCorpCust 取得客戶公司資料
        [HttpPost]
        public void GetCorpCust(int CompanyId = -1, int CsCustId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CorpCustManagement", "read");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetCorpCust(CompanyId, CsCustId, Status
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
        #region //AddModule 模組別資料新增
        [HttpPost]
        public void AddModule(int SystemId = -1, string ModuleCode = "", string ModuleName = "", string ThemeIcon = "")
        {
            try
            {
                WebApiLoginCheck("ModuleSetting", "add");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.AddModule(SystemId, ModuleCode, ModuleName, ThemeIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.AddFunction(ModuleId, FunctionCode, FunctionName, UrlTarget);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void AddFunctionDetail(int FunctionId = -1, string FnDetailCode = "", string FnDetailName = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "add");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.AddFunctionDetail(FunctionId, FnDetailCode, FnDetailName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddRole 角色資料新增
        [HttpPost]
        public void AddRole(string RoleName = "", string AdminStatus = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.AddRole(RoleName, AdminStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddCsCustomerDetail 會員客戶明細新增
        [HttpPost]
        public void AddCsCustomerDetail(int CompanyId = -1, string CustomerName = "", string CustomerEnglishName = "", string Customers = "")
        {
            try
            {
                WebApiLoginCheck("CorpCustManagement", "add");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.AddCsCustomerDetail(CompanyId, CustomerName, CustomerEnglishName, Customers);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateModule 模組別資料更新
        [HttpPost]
        public void UpdateModule(int ModuleId = -1, int SystemId = -1, string ModuleCode = "", string ModuleName = "", string ThemeIcon = "")
        {
            try
            {
                WebApiLoginCheck("ModuleSetting", "update");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateModule(ModuleId, SystemId, ModuleCode, ModuleName, ThemeIcon);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateModuleStatus(ModuleId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateModuleSort(SystemId, ModuleList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateFunction(FunctionId, ModuleId, FunctionCode, FunctionName, UrlTarget);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateFunctionStatus(FunctionId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateFunctionSort(ModuleId, FunctionList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void UpdateFunctionDetail(int FnDetailId = -1, int FunctionId = -1, string FnDetailCode = "", string FnDetailName = "")
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "update");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateFunctionDetail(FnDetailId, FunctionId, FnDetailCode, FnDetailName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void UpdateFunctionDetailStatus(int FnDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("FunctionSetting", "status-switch");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateFunctionDetailStatus(FnDetailId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateFunctionDetailSort(FunctionId, FunctionDetailList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRole 角色資料更新
        [HttpPost]
        public void UpdateRole(int RoleId = -1, string RoleName = "", string AdminStatus = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateRole(RoleId, RoleName, AdminStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoleStatus 角色狀態更新
        [HttpPost]
        public void UpdateRoleStatus(int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateRoleStatus(RoleId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoleUser 角色人員更新
        [HttpPost]
        public void UpdateRoleUser(int RoleId = -1, string UserList = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateRoleUser(RoleId, UserList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateRoleFunctionDetail 角色權限更新
        [HttpPost]
        public void UpdateRoleFunctionDetail(int RoleId = -1, int FnDetailId = -1, bool Checked = false, string SearchKey = "")
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateRoleFunctionDetail(RoleId, FnDetailId, Checked, SearchKey);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateUserFunctionDetail 使用者權限更新
        [HttpPost]
        public void UpdateUserFunctionDetail(int RoleId = -1, string UserList = "", bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateUserFunctionDetail(RoleId, UserList, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMemberCustomer 會員客戶資料更新
        [HttpPost]
        [Route("api/SYS/UpdateMemberCustomer")]
        public void UpdateMemberCustomer(string UserList = "", int CsCustId = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "update");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateMemberCustomer(UserList, CsCustId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCsCustomerDetail 會員客戶明細更新
        [HttpPost]
        public void UpdateCsCustomerDetail(int CompanyId = -1, int CsCustId = -1, string CustomerName = "", string CustomerEnglishName = "", string Customers = "")
        {
            try
            {
                WebApiLoginCheck("CorpCustManagement", "update");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateCsCustomerDetail(CompanyId, CsCustId, CustomerName, CustomerEnglishName, Customers);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateCsCustomerStatus 會員客戶狀態更新
        [HttpPost]
        public void UpdateCsCustomerStatus(int CsCustId = -1)
        {
            try
            {
                WebApiLoginCheck("CorpCustManagement", "status-switch");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateCsCustomerStatus(CsCustId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMemberStatus 會員狀態更新
        [HttpPost]
        public void UpdateMemberStatus(int MemberId = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "status-switch");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.UpdateMemberStatus(MemberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteRoleUser 角色會員刪除
        [HttpPost]
        public void DeleteRoleUser(int MemberId = -1, int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("AuthorityManagement", "authority-setting");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.DeleteRoleUser(MemberId, RoleId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMemberCustomer 會員客戶資料刪除
        [HttpPost]
        public void DeleteMemberCustomer(int MemberId = -1)
        {
            try
            {
                WebApiLoginCheck("MemberManagement", "update");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.DeleteMemberCustomer(MemberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteCsCustomer 會員客戶資料刪除
        [HttpPost]
        public void DeleteCsCustomer(int CsCustId = -1)
        {
            try
            {
                WebApiLoginCheck("CorpCustManagement", "delete");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.DeleteCsCustomer(CsCustId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteCsCustomerDetail 會員客戶明細資料刪除
        [HttpPost]
        public void DeleteCsCustomerDetail(int CsCustId = -1, int CustomerId = -1)
        {
            try
            {
                WebApiLoginCheck("CorpCustManagement", "delete");

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.DeleteCsCustomerDetail(CsCustId, CustomerId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //GetModule 取得模組別資料
        [HttpPost]
        [Route("api/SYS/GetModule")]
        public void GetModule(int ModuleId = -1, int SystemId = -1, string ModuleCode = "", string ModuleName = "", string Status = "", int[] CustomerIds = null
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "GetNotificationMail");
                #endregion

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetModule(ModuleId, SystemId, ModuleCode, ModuleName, Status
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

        #region //GetSystemMenu 取得系統選單資料
        [HttpPost]
        [Route("api/SYS/GetSystemMenu")]
        public void GetSystemMenu(int MemberId = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "GetNotificationMail");
                #endregion

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetSystemMenu(MemberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        [Route("api/SYS/GetModuleMenu")]
        public void GetModuleMenu(string SystemCode = "", int MemberId = -1, int[] CustomerIds = null)
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "GetNotificationMail");
                #endregion

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetModuleMenu(SystemCode, MemberId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetFunction 取得模組功能資料
        [HttpPost]
        [Route("api/SYS/GetFunction")]
        public void GetFunction(int FunctionId = -1, int ModuleId = -1, int SystemId = -1, string FunctionCode = "", string FunctionName = "", string Status = "", int[] CustomerIds = null
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "GetNotificationMail");
                #endregion

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetFunction(FunctionId, ModuleId, SystemId, FunctionCode, FunctionName, Status
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

        #region //GetAuthorityVerify 權限驗證
        [HttpPost]
        [Route("api/SYS/GetAuthorityVerify")]
        public void GetAuthorityVerify(int MemberId, int[] CustomerIds, string FunctionCode = "", string FnDetailCode = "")
        {
            try
            {
                #region //Api金鑰驗證
                //ApiKeyVerify(Company, SecretKey, "GetNotificationMail");
                #endregion

                #region //Request
                eipSystemSettingDA = new EipSystemSettingDA();
                dataRequest = eipSystemSettingDA.GetAuthorityVerify(MemberId, CustomerIds, FunctionCode, FnDetailCode);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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