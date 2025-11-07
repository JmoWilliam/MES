using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PMISDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class PmisBasicInformationController : WebController
    {
        private PmisBasicInformationDA pmisBasicInformationDA = new PmisBasicInformationDA();

        #region //View
        public ActionResult Authority()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult FileLevel()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TemplateManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetAuthority 取得專案權限資料
        [HttpPost]
        public void GetAuthority(int AuthorityId = -1, string AuthorityType = "", string AuthorityName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("Authority", "read");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetAuthority(AuthorityId, AuthorityType, AuthorityName, Status
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

        #region //GetFileLevel 取得檔案等級資料
        [HttpPost]
        public void GetFileLevel(int LevelId = -1, string LevelName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("FileLevel", "read,constrained-data");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetFileLevel(LevelId, LevelName, Status
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

        #region //GetTemplate 取得樣板資料
        [HttpPost]
        public void GetTemplate(int TemplateId = -1, string TemplateName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "read");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetTemplate(TemplateId, TemplateName, Status
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

        #region //GetTemplateCooperation 取得樣板協作人員資料
        [HttpPost]
        public void GetTemplateCooperation(int TemplateId = -1, int CompanyId = -1, int DepartmentId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "cooperation");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetTemplateCooperation(TemplateId, CompanyId, DepartmentId, SearchKey
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

        #region //GetTemplateTask 取得樣板任務資料
        [HttpPost]
        public void GetTemplateTask(int TemplateTaskId = -1, int TemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "read");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetTemplateTask(TemplateTaskId, TemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTemplateTaskTree 取得樣板任務資料(樹狀結構)
        [HttpPost]
        public void GetTemplateTaskTree(int TemplateId = -1, int ParentTaskId = -1, int OpenedLevel = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "read");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetTemplateTaskTree(TemplateId, ParentTaskId, OpenedLevel);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTemplateTaskGantt 取得樣板任務資料(甘特圖)
        [HttpPost]
        public void GetTemplateTaskGantt(int TemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "read");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetTemplateTaskGantt(TemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTemplateTaskUser 取得樣板任務成員資料
        [HttpPost]
        public void GetTemplateTaskUser(int TemplateTaskId = -1, int CompanyId = -1, int DepartmentId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.GetTemplateTaskUser(TemplateTaskId, CompanyId, DepartmentId, SearchKey
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
        #region //AddAuthority 專案權限資料新增
        [HttpPost]
        public void AddAuthority(string AuthorityType = "", string AuthorityCode = "", string AuthorityName = "")
        {
            try
            {
                WebApiLoginCheck("Authority", "add");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddAuthority(AuthorityType, AuthorityCode, AuthorityName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddFileLevel 檔案等級資料新增
        [HttpPost]
        public void AddFileLevel(string LevelCode = "", string LevelName = "")
        {
            try
            {
                WebApiLoginCheck("FileLevel", "add");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddFileLevel(LevelCode, LevelName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplate 樣板資料新增
        [HttpPost]
        public void AddTemplate(string TemplateName = "", string TemplateDesc = "", string WorkTimeStatus = "", string WorkTimeInterval = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "add");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddTemplate(TemplateName, TemplateDesc, WorkTimeStatus, WorkTimeInterval);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateCooperation 樣板協作人員資料新增
        [HttpPost]
        public void AddTemplateCooperation(int TemplateId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "cooperation");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddTemplateCooperation(TemplateId, Users);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateTask 樣板任務資料新增
        [HttpPost]
        public void AddTemplateTask(int TemplateId = -1, int ParentTaskId = -1, string TaskName = "", int Duration = -1, int ReplyFrequency = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-add");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddTemplateTask(TemplateId, ParentTaskId, TaskName, Duration, ReplyFrequency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateTaskLink 樣板任務連結資料新增
        [HttpPost]
        public void AddTemplateTaskLink(int SourceTaskId = -1, int TargetTaskId = -1, string LinkType = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddTemplateTaskLink(SourceTaskId, TargetTaskId, LinkType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTemplateTaskUser 樣板任務成員新增
        [HttpPost]
        public void AddTemplateTaskUser(int TemplateTaskId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.AddTemplateTaskUser(TemplateTaskId, Users);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateAuthority 專案權限資料更新
        [HttpPost]
        public void UpdateAuthority(int AuthorityId = -1, string AuthorityName = "")
        {
            try
            {
                WebApiLoginCheck("Authority", "update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateAuthority(AuthorityId, AuthorityName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateAuthorityStatus 專案權限狀態更新
        [HttpPost]
        public void UpdateAuthorityStatus(int AuthorityId = -1)
        {
            try
            {
                WebApiLoginCheck("Authority", "status-switch");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateAuthorityStatus(AuthorityId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateAuthoritySort 專案權限順序調整
        [HttpPost]
        public void UpdateAuthoritySort(string AuthorityType = "", string AuthorityList = "")
        {
            try
            {
                WebApiLoginCheck("Authority", "sort");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateAuthoritySort(AuthorityType, AuthorityList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFileLevel 檔案等級資料更新
        [HttpPost]
        public void UpdateFileLevel(int LevelId = -1, string LevelName = "")
        {
            try
            {
                WebApiLoginCheck("FileLevel", "update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateFileLevel(LevelId, LevelName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFileLevelStatus 檔案等級狀態更新
        [HttpPost]
        public void UpdateFileLevelStatus(int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("FileLevel", "status-switch");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateFileLevelStatus(LevelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateFileLevelSort 檔案等級順序調整
        [HttpPost]
        public void UpdateFileLevelSort(string FileLevelList = "")
        {
            try
            {
                WebApiLoginCheck("FileLevel", "sort");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateFileLevelSort(FileLevelList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplate 樣板資料更新
        [HttpPost]
        public void UpdateTemplate(int TemplateId = -1, string TemplateName = "", string TemplateDesc = "", string WorkTimeStatus = "", string WorkTimeInterval = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplate(TemplateId, TemplateName, TemplateDesc, WorkTimeStatus, WorkTimeInterval);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateStatus 樣板資料狀態更新
        [HttpPost]
        public void UpdateTemplateStatus(int TemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "status-switch");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateStatus(TemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateTask 樣板任務資料更新
        [HttpPost]
        public void UpdateTemplateTask(int TemplateTaskId = -1, string TaskName = "", int Duration = -1, int ReplyFrequency = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateTask(TemplateTaskId, TaskName, Duration, ReplyFrequency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateTaskSort 樣板任務順序調整
        [HttpPost]
        public void UpdateTemplateTaskSort(int TemplateTaskId = -1, string SortType = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-sort");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateTaskSort(TemplateTaskId, SortType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateTaskSchedule 樣板任務排程更新(單一任務)
        [HttpPost]
        public void UpdateTemplateTaskSchedule(int TemplateTaskId = -1, string StartDate = "", int Duration = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateTaskSchedule(TemplateTaskId, StartDate, Duration);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateAllTaskSchedule 樣板任務排程更新(全部任務)
        [HttpPost]
        public void UpdateTemplateAllTaskSchedule(string TaskData)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateAllTaskSchedule(TaskData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateTaskUserAgentSort 樣板任務成員代理順序更新
        [HttpPost]
        public void UpdateTemplateTaskUserAgentSort(int TemplateTaskId = -1, string TemplateTaskUserList = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateTaskUserAgentSort(TemplateTaskId, TemplateTaskUserList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateTaskUserAuthority 樣板任務成員權限更新
        [HttpPost]
        public void UpdateTemplateTaskUserAuthority(int TemplateTaskUserId = -1, int AuthorityId = -1, bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateTaskUserAuthority(TemplateTaskUserId, AuthorityId, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateAllTaskUserAuthority 批次更新任務成員權限
        public void UpdateTemplateAllTaskUserAuthority(int TemplateTaskUserId = -1, bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateAllTaskUserAuthority(TemplateTaskUserId, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemplateTaskUserLevel 樣板任務成員閱覽檔案等級更新
        [HttpPost]
        public void UpdateTemplateTaskUserLevel(int TemplateTaskUserId = -1, int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.UpdateTemplateTaskUserLevel(TemplateTaskUserId, LevelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteTemplate 樣板資料刪除
        [HttpPost]
        public void DeleteTemplate(int TemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "delete");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.DeleteTemplate(TemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateCooperation 樣板協作人員資料刪除
        [HttpPost]
        public void DeleteTemplateCooperation(int TemplateCooperationId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "cooperation");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.DeleteTemplateCooperation(TemplateCooperationId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateTask 樣板任務資料刪除
        [HttpPost]
        public void DeleteTemplateTask(int TemplateTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-delete");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.DeleteTemplateTask(TemplateTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateTaskLink 樣板任務連結資料刪除
        [HttpPost]
        public void DeleteTemplateTaskLink(int TemplateTaskLinkId = -1, int SourceTaskId = -1, int TargetTaskId = -1, string LinkType = "")
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "task-update");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.DeleteTemplateTaskLink(TemplateTaskLinkId, SourceTaskId, TargetTaskId, LinkType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteTemplateTaskUser 樣板任務成員刪除
        [HttpPost]
        public void DeleteTemplateTaskUser(int TemplateTaskUserId = -1)
        {
            try
            {
                WebApiLoginCheck("TemplateManagement", "member");

                #region //Request
                pmisBasicInformationDA = new PmisBasicInformationDA();
                dataRequest = pmisBasicInformationDA.DeleteTemplateTaskUser(TemplateTaskUserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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