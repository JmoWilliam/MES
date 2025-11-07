using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PMISDA;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace Business_Manager.Controllers
{
    public class TaskController : WebController
    {
        private ProjectDA projectDA = new ProjectDA();
        private TaskDA taskDA = new TaskDA();

        #region //View
        public ActionResult TaskManagement()
        {
            ViewLoginCheck();

            return View();
        }

        [HttpPost]
        [Route("api/ReplyInterfaceApi")]
        public ActionResult ReplyInterfaceApi(string SecretKey = "", string Company = "", int ProjectTaskId = -1, string UserNo = "")
        {
            return RedirectToAction("ReplyInterface", "api", new { SecretKey, Company, ProjectTaskId, UserNo });
        }

        [Route("api/ReplyInterface")]
        public ActionResult ReplyInterface(string SecretKey = "", string Company = "", int ProjectTaskId = -1, string UserNo = "")
        {
            return View();
        }
        #endregion

        #region //Get
        #region //GetTaskChecklist 取得任務清單資料
        [HttpPost]
        public void GetTaskChecklist(string Project = "", string TaskName = "", string TaskStatus = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetTaskChecklist(Project, TaskName, TaskStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPersonalTaskChecklist -- 取得個人任務清單資料-New
        [HttpPost]
        public void GetPersonalTaskChecklist(string TaskName = "", string TaskStatus = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetPersonalTaskChecklist(TaskName, TaskStatus);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTaskStructure 取得任務結構資料
        [HttpPost]
        public void GetTaskStructure(string Project = "", int ProjectTaskId = -1, string TaskName = "", string TaskStatus = "", int OpenedLevel = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetTaskStructure(Project, ProjectTaskId, TaskName, TaskStatus, OpenedLevel);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserProject 取得使用者專案資料
        [HttpPost]
        public void GetUserProject()
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetUserProject();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserTask 取得使用者任務資料
        [HttpPost]
        public void GetUserTask(int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetUserTask(ProjectId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserPersonalTask -- 取得使用者個人任務資料-New
        [HttpPost]
        public void GetUserPersonalTask(int PersonalTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetUserPersonalTask(PersonalTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetUserTaskCalendar --取得行事曆任務資料 -- Chia Yuan 2023.11.29
        [HttpPost]
        public void GetUserTaskCalendar(int ProjectId = -1, int ProjectTaskId = -1, string TaskType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetUserTaskCalendar(ProjectId, ProjectTaskId, TaskType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectTaskFile 取得使用者可存取檔案資料
        [HttpPost]
        public void GetProjectTaskFile(int ProjectId = -1, int TaskId = -1, string FileName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetProjectTaskFile(ProjectId, TaskId, FileName, OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetPersonalTaskFile 取得使用者可存取檔案資料-New
        [HttpPost]
        public void GetPersonalTaskFile(int PersonalTaskId = -1, string FileName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetPersonalTaskFile(PersonalTaskId, FileName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void GetProject(int ProjectTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProject(-1, -1, "", "", "", "", ProjectTaskId, "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                WebApiLoginCheck("TaskManagement", "read");

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

        #region //GetPersonalTask 取得個人任務資料
        [HttpPost]
        public void GetPersonalTask(int PersonalTaskId = -1, int UserId = -1, int CompanyId = -1, int DepartmentId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetPersonalTask(PersonalTaskId, UserId, CompanyId, DepartmentId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetLastReply 取得專案任務最後回報資料-New
        [HttpPost]
        public void GetLastReply(int TaskId = -1, string ReplyType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetLastReply(TaskId, ReplyType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                WebApiLoginCheck("TaskManagement", "read");

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

        #region //GetTaskGantt 取得任務資料(甘特圖)
        [HttpPost]
        public void GetTaskGantt(int TaskId = -1, string TaskType = "Project")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetTaskGantt(TaskId, TaskType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddTaskReply 任務回報內容新增-New
        [HttpPost]
        public void AddTaskReply(int TaskId = -1, string ReplyType = "", string ReplyStatus = "", string ActualStart = "", string ActualEnd = ""
            , int DeferredDuration = -1, string DurationUnit = "", int ReplyFile = -1, string ReplyContent = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.AddTaskReply(TaskId, ReplyType, ReplyStatus, ActualStart, ActualEnd, DeferredDuration, DurationUnit, ReplyFile, ReplyContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddProjectTaskFile 任務檔案資料新增
        [HttpPost]
        public void AddProjectTaskFile(int TaskId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddProjectTaskFile(TaskId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPersonalTaskFile -- 個人任務檔案資料新增
        [HttpPost]
        public void AddPersonalTaskFile(int PersonalTaskId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.AddPersonalTaskFile(PersonalTaskId, FileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPersonalTask -- 個人任務新增-New
        [HttpPost]
        public void AddPersonalTask(string TaskName = "", string TaskDesc = "", string PlannedStart = "", int PlannedDuration = -1, int ReplyFrequency = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "add");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.AddPersonalTask(TaskName, TaskDesc, PlannedStart, PlannedDuration, ReplyFrequency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddPersonalTaskByCopy -- 個人任務複製-New
        [HttpPost]
        public void AddPersonalTaskByCopy(int PersonalTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "add");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.AddPersonalTaskByCopy(PersonalTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateProjectAllTaskEstimate 專案任務排程預估更新(全部任務)
        [HttpPost]
        public void UpdateProjectAllTaskEstimate(int ProjectId = -1, int ProjectTaskId = -1, string TaskData = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.UpdateProjectAllTaskEstimate(ProjectId, ProjectTaskId, TaskData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePersonalTask -- 專案任務資料更新-New
        [HttpPost]
        public void UpdatePersonalTask(int PersonalTaskId = -1, string TaskName = "", string TaskDesc = ""
            , string PlannedStart = "", int PlannedDuration = -1, int ReplyFrequency = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.UpdatePersonalTask(PersonalTaskId, TaskName, TaskDesc
                    , PlannedStart, PlannedDuration, ReplyFrequency);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePersonalTaskFileLevel 個人任務檔案資閱覽檔案等級更新-New
        [HttpPost]
        public void UpdatePersonalTaskFileLevel(int PersonalTaskFileId = -1, int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdatePersonalTaskFileLevel(PersonalTaskFileId, LevelId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdatePersonalTaskFileType
        [HttpPost]
        public void UpdatePersonalTaskFileType(int PersonalTaskFileId = -1, string FileType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdatePersonalTaskFileType(PersonalTaskFileId, FileType);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteProjectTask 專案任務資料刪除
        [HttpPost]
        public void DeletePersonalTask(int PersonalTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = projectDA.DeletePersonalTask(PersonalTaskId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeletePersonalTaskFile 個人任務檔案刪除
        [HttpPost]
        public void DeletePersonalTaskFile(int PersonalTaskFileId = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.DeletePersonalTaskFile(PersonalTaskFileId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteReplyFile 回報檔案刪除
        [HttpPost]
        public void DeleteReplyFile(int key = -1)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.DeleteReplyFile(key);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //SetProjectTaskAuthority 設定專案任務獨立權限
        [HttpPost]
        public void SetProjectTaskAuthority(int ProjectTaskId)
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");

                Session["ProjectTaskId"] = ProjectTaskId;

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
                dataRequest = projectDA.GetTaskUserAuthority(ProjectTaskId, "Project", userNo, "");
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTaskUserAuthority -- 取得任務成員權限資料(前端)-New
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

        #region //CheckTaskAuthority 檢查任務權限
        private void CheckTaskAuthority(string authorityCode, int projectTaskId = -1)
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
                case "task-file-download":
                    break;
                default:
                    projectTaskId = Convert.ToInt32(Session["ProjectTaskId"]);
                    break;
            }

            dataRequest = projectDA.GetTaskUserAuthority(projectTaskId, "Project", userNo, authorityCode);
            jsonResponse = BaseHelper.DAResponse(dataRequest);

            if (jsonResponse["status"].ToString() == "success")
            {
                if (jsonResponse["result"].Count() <= 0)
                {
                    throw new SystemException("無此任務的【" + authorityCode + "】權限");
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

        #region //DownloadTaskFile 下載任務檔案(停用)
        //public ActionResult DownloadTaskFile(int FileId = -1, int TaskId = -1)
        //{
        //    try
        //    {
        //        WebApiLoginCheck("TaskManagement", "read");
        //        CheckTaskAuthority("task-file-download", TaskId);

        //        return RedirectToAction("GetFile", "Web", new { fileId = FileId });
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

        //        Response.ContentType = "application/json";
        //        Response.Write(jsonResponse.ToString(Formatting.None));

        //        return new EmptyResult();
        //    }
        //}
        #endregion

        #region //DownloadTaskFile 下載個人任務檔案-New
        public ActionResult DownloadTaskFile(int FileId = -1, int TaskId = -1, string TaskType = "")
        {
            try
            {
                WebApiLoginCheck("TaskManagement", "read");
                CheckTaskAuthority(TaskId, TaskType, "task-file-download");

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
                Response.Write(jsonResponse.ToString(Formatting.None));

                return new EmptyResult();
            }
        }
        #endregion
        #endregion

        #region //Api
        #region //Get
        #region //GetProjectApi 取得專案資料
        [HttpPost]
        public void GetProjectApi(int ProjectTaskId = -1
            , string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProject(-1, -1, "", "", "", "", ProjectTaskId, "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTaskGanttApi 取得任務資料(甘特圖)
        [HttpPost]
        public void GetTaskGanttApi(int ProjectTaskId = -1
            , string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.GetTaskGantt(ProjectTaskId, "Project");
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateProjectAllTaskEstimateApi 專案任務排程預估更新(全部任務)
        [HttpPost]
        public void UpdateProjectAllTaskEstimateApi(int ProjectId = -1, int ProjectTaskId = -1, string TaskData = ""
            , string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                taskDA = new TaskDA();
                dataRequest = taskDA.UpdateProjectAllTaskEstimate(ProjectId, ProjectTaskId, TaskData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetReplyTaskApi 取得可回報任務列表
        [HttpPost]
        [Route("api/GetReplyTaskApi")]
        public void GetReplyTaskApi(int ProjectId = -1, int CompanyId = -1, int DepartmentId = -1, string TaskUserNos = ""
            , string Company = "", string SecretKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetReplyTaskApi(ProjectId, CompanyId, DepartmentId, TaskUserNos
                    , "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectApi 取得專案資料
        [HttpPost]
        [Route("api/GetProjectApi")]
        public void GetProjectApi(int ProjectId = -1, int ProjectTaskId = -1, string ProjectUserNos = ""
            , string Company = "", string SecretKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectApi(ProjectId, ProjectTaskId, ProjectUserNos, "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectTaskApi 取得專案任務資料
        [HttpPost]
        [Route("api/GetProjectTaskApi")]
        public void GetProjectTaskApi(int ProjectTaskId = -1, string TaskUserNos = ""
            , string Company = "", string SecretKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.GetProjectTaskApi(ProjectTaskId, TaskUserNos
                    , "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateTaskReply 取得專案任務資料
        [HttpPost]
        [Route("api/UpdateTaskReply")]
        public void UpdateTaskReply(int ProjectTaskId = -1, string ReplyType = "", string ReplyStatus = "", string ActualStart = "", string ActualEnd = ""
            , int DeferredDuration = -1, string DurationUnit = "", int ReplyFile = -1, string ReplyContent = "", string UserNo = ""
            , string Company = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "RequestPmisFunction");
                #endregion

                #region //Request
                projectDA = new ProjectDA();
                dataRequest = projectDA.UpdateTaskReply(ProjectTaskId, ReplyType, ReplyStatus, ActualStart, ActualEnd
                    , DeferredDuration, DurationUnit, ReplyFile, ReplyContent, UserNo);
                #endregion

                #region //Response
                jsonResponse = JObject.Parse(dataRequest);
                if (jsonResponse["status"].ToString() == "errorForDA")
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = jsonResponse["msg"].ToString()
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = jsonResponse["msg"].ToString(),
                        result = Regex.Replace(Request.Url.AbsoluteUri, "api/.*|Task/.*", string.Format("api/ReplyInterface?SecretKey={0}&Company={1}&ProjectTaskId={2}&UserNo={3}", SecretKey, Company, ProjectTaskId, UserNo))
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

                logger.Error(e.Message);
            }
            Response.Write(jsonResponse.ToString());

            //TempData["data"] = (ProjectTaskId, UserNo);
            //return RedirectToAction("ReplyInterface", "Task", new { ProjectTaskId, UserNo });
        }
        #endregion
        #endregion
        #endregion
    }
}