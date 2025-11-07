using Helpers;
using PMISDA;
using System;
using System.Web;
using System.Linq;
using System.Web.Mvc;
using Newtonsoft.Json.Linq;

namespace Business_Manager.Controllers
{
    public class ProjectController : WebController
    {
        private ProjectDA projectDA = new ProjectDA();
        private TaskDA taskDA = new TaskDA();

        #region //View
        public ActionResult ProjectManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProjectExample()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TaskExample()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetProjectNo 取得專案代碼
        [HttpPost]
        public void GetProjectNo(string CompanyType = "", string ProjectType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectNo(CompanyType, ProjectType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProject 取得專案資料
        [HttpPost]
        public void GetProject(int ProjectId = -1, int DepartmentId = -1, string ProjectNo = "", string ProjectName = ""
            , string ProjectAttribute = "", string ProjectStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProject(ProjectId, DepartmentId, ProjectNo, ProjectName
                    , ProjectAttribute, ProjectStatus, -1
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

        #region //GetProjectUser 取得專案成員資料
        [HttpPost]
        public void GetProjectUser(int ProjectId = -1, int CompanyId = -1, int DepartmentId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-manage");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectUser(ProjectId, CompanyId, DepartmentId, SearchKey
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

        #region //GetProjectFile 取得專案檔案資料
        [HttpPost]
        public void GetProjectFile(int ProjectFileId = -1, int ProjectId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("file-manage");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectFile(ProjectFileId, ProjectId
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

        #region //GetProjectTask 取得專案任務資料
        [HttpPost]
        public void GetProjectTask(int ProjectTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-manage");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectTask(ProjectTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectTaskTree 取得專案任務資料(樹狀結構)
        [HttpPost]
        public void GetProjectTaskTree(int ProjectId = -1, int ParentTaskId = -1, int OpenedLevel = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-manage");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectTaskTree(ProjectId, ParentTaskId, OpenedLevel);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectTaskGantt 取得專案任務資料(甘特圖)
        [HttpPost]
        public void GetProjectTaskGantt(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-manage");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectTaskGantt(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectTaskUser 取得專案任務成員資料(停用)
        //[HttpPost]
        //public void GetProjectTaskUser(int ProjectTaskId = -1, int CompanyId = -1, int DepartmentId = -1, string SearchKey = ""
        //    , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        //{
        //    try
        //    {
        //        WebApiLoginCheck("ProjectManagement", "read");
        //        CheckProjectAuthority("task-update");
        //        #region //Request
        //        projectDA = new ProjectDA();
        //        dataRequest = projectDA.GetProjectTaskUser(ProjectTaskId, CompanyId, DepartmentId, SearchKey
        //            , OrderBy, PageIndex, PageSize);
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

        #region //GetTaskUser 取得任務成員資料-New (需傳入任務類型)
        [HttpPost]
        public void GetTaskUser(int TaskId = -1, int CompanyId = -1, int DepartmentId = -1, string SearchKey = "", string TaskType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetTaskUser(TaskId, CompanyId, DepartmentId, SearchKey, TaskType
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

        #region //GetTaskReply 取得任務回報內容資料
        [HttpPost]
        public void GetTaskReply(int TaskId = -1, string ReplyType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetTaskReply(TaskId, ReplyType
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

        #region //GetProjectTaskFile 取得專案任務檔案
        [HttpPost]
        public void GetProjectTaskFile(int ProjectTaskFileId = -1, int ProjectTaskId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectTaskFile(ProjectTaskFileId, ProjectTaskId
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
        #region //AddProject 專案資料新增
        [HttpPost]
        public void AddProject(int DepartmentId = -1, string CompanyType = "", string ProjectType = "", string ProjectNo = ""
            , string ProjectName = "", string ProjectDesc = "", string ProjectAttribute = "", string WorkTimeStatus = ""
            , string WorkTimeInterval = "", string CustomerRemark = "", string ProjectRemark = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProject(DepartmentId, CompanyType, ProjectType, ProjectNo, ProjectName
                    , ProjectDesc, ProjectAttribute, WorkTimeStatus, WorkTimeInterval, CustomerRemark, ProjectRemark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectUser 專案成員資料新增
        [HttpPost]
        public void AddProjectUser(int ProjectId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectUser(ProjectId, Users);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectFile 專案檔案資料新增
        [HttpPost]
        public void AddProjectFile(int ProjectId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("file-upload");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectFile(ProjectId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTask 專案任務資料新增
        [HttpPost]
        public void AddProjectTask(int ProjectId = -1, int ParentTaskId = -1, string TaskName = "", string TaskDesc = ""
            , string PlannedStart = "", int PlannedDuration = -1, int ReplyFrequency = -1, bool RequiredFile = false)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTask(ProjectId, ParentTaskId, TaskName, TaskDesc
                    , PlannedStart, PlannedDuration, ReplyFrequency, RequiredFile);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTaskByCopy 新增複製任務
        [HttpPost]
        public void AddProjectTaskByCopy(int ProjectTaskId)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTaskByCopy(ProjectTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTaskByTemplate 專案任務資料新增(樣板)
        [HttpPost]
        public void AddProjectTaskByTemplate(int ProjectId = -1, int ParentTaskId = -1, int TemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTaskByTemplate(ProjectId, ParentTaskId, TemplateId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTaskLink 專案任務連結資料新增
        [HttpPost]
        public void AddProjectTaskLink(int SourceTaskId = -1, int TargetTaskId = -1, string LinkType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTaskLink(SourceTaskId, TargetTaskId, LinkType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTaskUser 專案任務成員新增
        [HttpPost]
        public void AddProjectTaskUser(int ProjectTaskId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTaskUser(ProjectTaskId, Users);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddTaskUser 任務成員新增-New (需傳入任務類型)
        [HttpPost]
        public void AddTaskUser(int TaskId = -1, string TaskType = "", string Users = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddTaskUser(TaskId, TaskType, Users);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTaskFile 專案任務檔案資料新增
        [HttpPost]
        public void AddProjectTaskFile(int ProjectTaskId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTaskFile(ProjectTaskId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectMamo 建立專案MAMO資料
        public void AddProjectMamo(int ProjectId)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "add");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectMamo(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateProject 專案資料更新
        [HttpPost]
        public void UpdateProject(int ProjectId = -1, int DepartmentId = -1, string ProjectName = "", string ProjectDesc = ""
            , string ProjectAttribute = "", string WorkTimeStatus = "", string WorkTimeInterval = "", string CustomerRemark = ""
            , string ProjectRemark = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("project-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProject(ProjectId, DepartmentId, ProjectName, ProjectDesc
                    , ProjectAttribute, WorkTimeStatus, WorkTimeInterval, CustomerRemark, ProjectRemark);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectStartUp 專案啟動
        [HttpPost]
        public void UpdateProjectStartUp(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "project-start-up");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectStartUp(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectReset 專案重置
        [HttpPost]
        public void UpdateProjectReset(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "project-reset");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectReset(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateAllTaskRefresh 專案重整
        [HttpPost]
        public void UpdateAllTaskRefresh(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("project-refresh");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateAllTaskRefresh(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectUserAgentSort 專案成員代理順序更新
        [HttpPost]
        public void UpdateProjectUserAgentSort(int ProjectId = -1, string ProjectUserList = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectUserAgentSort(ProjectId, ProjectUserList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectUserAuthority 專案成員權限更新
        [HttpPost]
        public void UpdateProjectUserAuthority(int ProjectUserId = -1, int AuthorityId = -1, bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectUserAuthority(ProjectUserId, AuthorityId, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectAllUserAuthority
        [HttpPost]
        public void UpdateProjectAllUserAuthority(int ProjectUserId = -1, bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectAllUserAuthority(ProjectUserId, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectUserLevel 專案成員閱覽檔案等級更新
        [HttpPost]
        public void UpdateProjectUserLevel(int ProjectUserId = -1, int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectUserLevel(ProjectUserId, LevelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectFileLevel 專案檔案資閱覽檔案等級更新
        [HttpPost]
        public void UpdateProjectFileLevel(int ProjectFileId = -1, int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("file-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectFileLevel(ProjectFileId, LevelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectFileFileType 專案檔案檔案類型更新
        [HttpPost]
        public void UpdateProjectFileFileType(int ProjectFileId = -1, string FileType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("file-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectFileFileType(ProjectFileId, FileType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectTask 專案任務資料更新
        [HttpPost]
        public void UpdateProjectTask(int ProjectTaskId = -1, string TaskName = "", string TaskDesc = "", int ReplyFrequency = -1, bool RequiredFile = false
            , string PlannedStart = "", int PlannedDuration = -1, string EstimateStart = "", int EstimateDuration = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectTask(ProjectTaskId, TaskName, TaskDesc, ReplyFrequency, RequiredFile
                    , PlannedStart, PlannedDuration, EstimateStart, EstimateDuration);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectTaskSort 專案任務順序調整
        [HttpPost]
        public void UpdateProjectTaskSort(int ProjectTaskId = -1, string SortType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectTaskSort(ProjectTaskId, SortType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectTaskSchedule 專案任務排程更新(單一任務)
        [HttpPost]
        public void UpdateProjectTaskSchedule(int ProjectTaskId = -1, string PlannedStart = "", string PlannedEnd = "", int PlannedDuration = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectTaskSchedule(ProjectTaskId, PlannedStart, PlannedEnd, PlannedDuration);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectAllTaskSchedule 專案任務排程更新(全部任務)
        [HttpPost]
        public void UpdateProjectAllTaskSchedule(int ProjectId = -1, string TaskData = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectAllTaskSchedule(ProjectId, TaskData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTemporaryProjectTask 插入、新增子任務排成更新(僅包含位移任務)
        [HttpPost]
        public void UpdateTemporaryProjectTask(int ProjectId = -1, int ProjectTaskId = -1, string TaskData = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateTemporaryProjectTask(ProjectId, ProjectTaskId, TaskData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTaskUserAgentSort 任務成員代理順序更新-New (需傳入任務類型)
        [HttpPost]
        public void UpdateTaskUserAgentSort(int TaskId = -1, string TaskType = "", string TaskUserList = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateTaskUserAgentSort(TaskId, TaskType, TaskUserList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTaskUserAuthority 任務成員權限更新-New (需傳入任務類型)
        [HttpPost]
        public void UpdateTaskUserAuthority(int TaskUserId = -1, int AuthorityId = -1, string TaskType = "", bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateTaskUserAuthority(TaskUserId, AuthorityId, TaskType, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateAllTaskUserAuthority 批次任務成員權限更新-New (需傳入任務類型)
        public void UpdateAllTaskUserAuthority(int TaskUserId = -1, string TaskType = "", bool Checked = false)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateAllTaskUserAuthority(TaskUserId, TaskType, Checked);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTaskUserLevel 任務成員閱覽檔案等級更新-New (需傳入任務類型)
        [HttpPost]
        public void UpdateTaskUserLevel(int TaskUserId = -1, int LevelId = -1, string TaskType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateTaskUserLevel(TaskUserId, LevelId, TaskType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectTaskFileLevel 專案任務檔案資閱覽檔案等級更新
        [HttpPost]
        public void UpdateProjectTaskFileLevel(int ProjectTaskFileId = -1, int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectTaskFileLevel(ProjectTaskFileId, LevelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectTaskFileType 專案任務檔案檔案類型更新
        [HttpPost]
        public void UpdateProjectTaskFileType(int ProjectTaskFileId = -1, string FileType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateProjectTaskFileType(ProjectTaskFileId, FileType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteProject 專案資料刪除
        [HttpPost]
        public void DeleteProject(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("project-delete", ProjectId);

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteProject(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProjectUser 專案成員資料刪除
        [HttpPost]
        public void DeleteProjectUser(int ProjectUserId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-delete");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteProjectUser(ProjectUserId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProjectFile 專案檔案刪除
        [HttpPost]
        public void DeleteProjectFile(int ProjectFileId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("file-delete");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteProjectFile(ProjectFileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProjectTask 專案任務資料刪除
        [HttpPost]
        public void DeleteProjectTask(int ProjectTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("user-delete");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteProjectTask(ProjectTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProjectTaskLink 專案任務連結資料刪除
        [HttpPost]
        public void DeleteProjectTaskLink(int ProjectTaskLinkId = -1, int SourceTaskId = -1, int TargetTaskId = -1, string LinkType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteProjectTaskLink(ProjectTaskLinkId, SourceTaskId, TargetTaskId, LinkType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProjectTaskUser 專案任務成員刪除 (停用)
        //[HttpPost]
        //public void DeleteProjectTaskUser(int TaskUserId = -1, string TaskType = "")
        //{
        //    try
        //    {
        //        WebApiLoginCheck("ProjectManagement", "read");
        //        CheckProjectAuthority("task-update");

        //        #region //Request
        //        projectDA = new ProjectDA();
        //        dataRequest = projectDA.DeleteTaskUser(TaskUserId, TaskType);
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

        #region //DeleteTaskUser 任務成員刪除-New (需傳入任務類型)
        [HttpPost]
        public void DeleteTaskUser(int TaskId = -1, int TaskUserId = -1, string TaskType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "delete");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteTaskUser(TaskId, TaskUserId, TaskType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteProjectTaskFile 專案任務檔案刪除
        [HttpPost]
        public void DeleteProjectTaskFile(int ProjectTaskFileId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("task-update");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeleteProjectTaskFile(ProjectTaskFileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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

        #region //Other
        #region //SetProjectAuthority 設定專案獨立權限(前端頁面)
        [HttpPost]
        public void SetProjectAuthority(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");

                Session["ProjectId"] = ProjectId;

                string userNo = "";
                #region //Cookie
                HttpCookie Login = Request.Cookies.Get("Login");

                if (Login != null)
                {
                    userNo = Login.Value;
                }
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectUserAuthority(ProjectId, userNo, "");
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTaskUserAuthority -- 取得任務成員權限資料(前端)-New (需傳入任務類型)
        [HttpPost]
        public void GetTaskUserAuthority(int TaskId = -1, string TaskType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                string userNo = "";
                #region //Cookie
                HttpCookie Login = Request.Cookies.Get("Login");

                if (Login != null)
                {
                    userNo = Login.Value;
                }
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetTaskUserAuthority(TaskId, TaskType, userNo, "");
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //CheckProjectAuthority 檢查專案權限
        private void CheckProjectAuthority(string authorityCode, int projectId = -1)
        {
            string userNo = "";
            #region //Cookie
            HttpCookie Login = Request.Cookies.Get("Login");

            if (Login != null)
            {
                userNo = Login.Value;
            }
            #endregion

            switch (authorityCode)
            {
                case "project-delete":
                case "file-download":
                    break;
                default:
                    projectId = Convert.ToInt32(Session["ProjectId"]);
                    break;
            }

            dataRequest = projectDA.GetProjectUserAuthority(projectId, userNo, authorityCode);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    throw new SystemException("無此專案的【" + authorityCode + "】權限");
                }
            }
            else
            {
                throw new SystemException(jsonResponse["msg"].ToString());
            }
        }
        #endregion

        #region //CheckTaskAuthority 檢查個人任務權限-New
        private void CheckTaskAuthority(int TaskId = -1, string TaskType = "", string AuthorityCode = "")
        {
            string userNo = "";
            #region //Cookie
            HttpCookie Login = Request.Cookies.Get("Login");

            if (Login != null)
            {
                userNo = Login.Value;
            }
            #endregion

            dataRequest = projectDA.GetTaskUserAuthority(TaskId, TaskType, userNo, AuthorityCode);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    throw new SystemException("無此任務的【" + AuthorityCode + "】權限");
                }
            }
            else
            {
                throw new SystemException(jsonResponse["msg"].ToString());
            }
        }
        #endregion

        #region //DownloadProjectFile 下載專案檔案
        public ActionResult DownloadProjectFile(int FileId, int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectManagement", "read");
                CheckProjectAuthority("file-download", ProjectId);

                return RedirectToAction("GetFile", "Web", new { fileId = FileId });
            }
            catch (Exception e)
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
        #endregion
    }
}