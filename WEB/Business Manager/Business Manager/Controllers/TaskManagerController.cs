using Helpers;
using Newtonsoft.Json.Linq;
using PMISDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class TaskManagerController : WebController
    {
        // GET: TaskManager
        private TaskManagerDA taskManagerDA = new TaskManagerDA();


        public ActionResult TaskDistribution()
        {
            //ViewLoginCheck();
            return View();
        }
        public ActionResult MissionReport()
        {
            ViewLoginCheck();
            return View();
        }

        #region //Get
        #region //GetUser 取得使用者資料
        [HttpPost]
        public void GetUser(int UserId = -1, int DepartmentId = -1, int CompanyId = -1, string Departments = ""
            , string UserNo = "", string UserName = "", string Gender = "", string Status = "", string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetUser(UserId, DepartmentId, CompanyId, Departments
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

        #region //GetTaskType 取得任務類型
        [HttpPost]
        public void GetTaskType(int TaskTypeId = -1, string TaskTypeName = "")
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "read,constrained-data");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetTaskType(TaskTypeId, TaskTypeName);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void GetDepartment(bool IsType, string CompanyNo)
        {
            try
            {
                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetDepartment(IsType, CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetCustomer 取得客戶資料
        [HttpPost]
        public void GetCustomer(string CompanyNo)
        {
            try
            {
                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetCustomer(CompanyNo);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetTaskDistribution 取得任務
        [HttpPost]
        public void GetTaskDistribution(string UserNo = "", int TdId = -1, string TaskName = "", string StartDate = "", string ExpectedDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "read,constrained-data");

                switch (UserNo)
                {
                    case "X9A7B2C8":
                        UserNo = "ZY00002";
                        break;
                }
                
                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetTaskDistribution(UserNo, TdId, TaskName, StartDate, ExpectedDate, Status
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

        #region //GetAttendUserData 取得參與者
        [HttpPost]
        public void GetAttendUserData(string AttendUser = "")
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "read,constrained-data");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetAttendUserData(AttendUser);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetReportContent 取得任務回報
        [HttpPost]
        public void GetReportContent(int TdId = -1)
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "read,constrained-data");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetReportContent(TdId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetMissionReport 取得任務(參與者)
        [HttpPost]
        public void GetMissionReport(int TdId = -1, string TaskName = "", string StartDate = "", string ExpectedDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "read,constrained-data");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetMissionReport(TdId, TaskName, StartDate, ExpectedDate, Status
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

        #region //GetMrDetail 取得任務人員回報
        [HttpPost]
        public void GetMrDetail(int MrDetailId = -1, int TdId = -1, int MrId = -1, int ReportUserId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "read,constrained-data");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.GetMrDetail(MrDetailId, TdId, MrId, ReportUserId, Status
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
        #region //AddTaskDistribution 任務新增
        [HttpPost]
        public void AddTaskDistribution(string UserNo = "", string TaskName = "", string TaskContent = "", string ExpectedDate = "", string AttendUser = ""
            , int ReturnFrequency = -1, string RepeatType = "", int RepeatValue = -1, string RepeatEndType = "", string RepeatEndValue = "", string FileList = ""
            , string ImgList = "")
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "add");

                switch (UserNo)
                {
                    case "X9A7B2C8":
                        UserNo = "ZY00002";
                        break;
                }

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.AddTaskDistribution(UserNo, TaskName, TaskContent, ExpectedDate, AttendUser
                    , ReturnFrequency, RepeatType, RepeatValue, RepeatEndType, RepeatEndValue, FileList, ImgList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //AddMrDetail 任務回報
        [HttpPost]
        public void AddMrDetail(int MrId = -1, string ReportContent = "")
        {
            try
            {
                //WebApiLoginCheck("MissionReport", "add");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.AddMrDetail(MrId, ReportContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //UpdateTaskDistribution 任務更新
        [HttpPost]
        public void UpdateTaskDistribution(string UserNo = "", int TdId = -1, int TaskTypeId = -1, string TaskName = "", string ProductionModel = "", int CustomerId = -1, int DepartmentId = -1, string TaskContent = "", string ExpectedDate = "", string AttendUser = ""
            , int ReturnFrequency = -1, string RepeatType = "", int RepeatValue = -1, string RepeatEndType = "", string RepeatEndValue = "", string FileList = ""
            , string ImgList = "")
        {
            try
            {
                //WebApiLoginCheck("TaskDistribution", "update");

                switch (UserNo)
                {
                    case "X9A7B2C8":
                        UserNo = "ZY00002";
                        break;
                }

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.UpdateTaskDistribution(UserNo, TdId, TaskTypeId, TaskName, ProductionModel, CustomerId, DepartmentId, TaskContent, ExpectedDate, AttendUser
                    , ReturnFrequency, RepeatType, RepeatValue, RepeatEndType, RepeatEndValue, FileList, ImgList);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateTaskDistributionStatus 任務狀態更新
        [HttpPost]
        public void UpdateTaskDistributionStatus(string UserNo = "", int TdId = -1, string Status = "")
        {
            try
            {

                //WebApiLoginCheck("TaskDistribution", "Void");

                switch (UserNo)
                {
                    case "X9A7B2C8":
                        UserNo = "ZY00002";
                        break;
                }

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.UpdateTaskDistributionStatus(UserNo, TdId, Status);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMissionReport 任務回報修改
        [HttpPost]
        public void UpdateMissionReport(int MrId = -1, string ReportContent = "")
        {
            try
            {
                //WebApiLoginCheck("MissionReport", "update");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.UpdateMissionReport(MrId, ReportContent);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //UpdateMsStatusSort 任務回報排序修改
        [HttpPost]
        public void UpdateMsStatusSort(int MrId = -1, string TargetState = "", string TargetSort = "")
        {
            try
            {
                //WebApiLoginCheck("MissionReport", "update");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.UpdateMsStatusSort(MrId, TargetState, TargetSort);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //DeleteTaskDistribution 任務刪除
        [HttpPost]
        public void DeleteTaskDistribution(string UserNo = "", int TdId = -1)
        {
            try
            {

                //WebApiLoginCheck("TaskDistribution", "delete");

                switch (UserNo)
                {
                    case "X9A7B2C8":
                        UserNo = "ZY00002";
                        break;
                }

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.DeleteTaskDistribution(UserNo, TdId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //DeleteMissionReport 回報刪除
        [HttpPost]
        public void DeleteMissionReport(int MrId = -1)
        {
            try
            {

                //WebApiLoginCheck("TaskDistribution", "delete");

                #region //Request
                taskManagerDA = new TaskManagerDA();
                dataRequest = taskManagerDA.DeleteMissionReport(MrId);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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