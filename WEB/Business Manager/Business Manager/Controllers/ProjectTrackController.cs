using Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PMISDA;
using Newtonsoft.Json.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System.IO;
using BASDA;

namespace Business_Manager.Controllers
{
    public class ProjectTrackController : WebController
    {
        private ProjectTrackDA projectTrackDA = new ProjectTrackDA();
        private ProjectDA projectDA = new ProjectDA();
        private TaskDA taskDA = new TaskDA();

        #region //View
        public ActionResult ProjectTrackManagement()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult TimelineDiagram()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult TimelineDiagram_bak()
        {
            return View();
        }
        #endregion

        public string GetDateTimeInfo(string dtString)
        {
            double.TryParse(dtString, out double duration);
            //DateTime.TryParse(task["TaskStart"]?.ToString(), out DateTime date1);
            //DateTime.TryParse(task["TaskEnd"]?.ToString(), out DateTime date2);
            //TimeSpan diff1 = date2.Subtract(date1);
            TimeSpan diff2 = TimeSpan.FromMinutes(duration);

            string durationInfo = string.Empty;
            if (duration >= 1440) //大於1天
            {
                durationInfo = string.Format("{0:D1}天 {1:D1}時 {2:D1}分", diff2.Days, diff2.Hours, diff2.Minutes);
            }
            else if (duration >= 60) //大於1小時
            {
                durationInfo = string.Format("{0:D1}時 {1:D1}分", diff2.Hours, diff2.Minutes);
            }
            else if (duration > 0) //大於1小時
            {
                durationInfo = string.Format("{0:D1}分", diff2.Minutes);
            }
            else {
                durationInfo = "-";
            }
            return durationInfo;
        }

        #region //Get
        #region //GetProjectTaskTree 取得專案任務資料(樹狀結構)
        [HttpPost]
        public void GetProjectTaskTree(int ProjectId = -1, string ProjectUserId = "", string TaskUserId = "", string ProjectMaster = "", int CompanyId = -1, int DepartmentId = -1, string Project = "", string TaskName = ""
            , string ProjectStatus = "", string TaskStatus = "", string ReplyStatus = "", string TaskStartDate = "", string TaskEndDate = ""
            , int OpenedLevel = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request authority
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityVerify("ProjectTrackManagement", "constrained-data");
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    if (jsonResponse["result"].Count() == 0)
                    {
                        CompanyId = Convert.ToInt32(Session["UserCompany"]);
                        DepartmentId = Convert.ToInt32(Session["UserDepartmentId"]);
                    }
                }
                #endregion

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectTaskTree(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate
                    , OpenedLevel
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

        #region //GetProjectTaskGantt 取得專案任務資料(甘特圖)
        [HttpPost]
        public void GetProjectTaskGantt(int ProjectId = -1, string ProjectUserId = "", string TaskUserId = "", string ProjectMaster = "", int CompanyId = -1, int DepartmentId = -1, string Project = "", string TaskName = ""
            , string ProjectStatus = "", string TaskStatus = "", string ReplyStatus = "", string TaskStartDate = "", string TaskEndDate = ""
            , string TreeGroupBy = "")
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request authority
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityVerify("ProjectTrackManagement", "constrained-data");
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    if (jsonResponse["result"].Count() == 0)
                    {
                        CompanyId = Convert.ToInt32(Session["UserCompany"]);
                        DepartmentId = Convert.ToInt32(Session["UserDepartmentId"]);
                    }
                }
                #endregion

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectTaskGantt(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate
                    , TreeGroupBy);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProjectTaskTimeline 取得專案任務資料(時間軸圖)
        [HttpPost]
        public void GetProjectTaskTimeline(int ProjectId = -1, string ProjectUserId = "", string TaskUserId = "", string ProjectMaster = "", int CompanyId = -1, int DepartmentId = -1, string Project = "", string TaskName = ""
            , string ProjectStatus = "", string TaskStatus = "", string ReplyStatus = "", string TaskStartDate = "", string TaskEndDate = ""
            , string DurationSplit = "")
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request authority
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityVerify("ProjectTrackManagement", "constrained-data");
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    if (jsonResponse["result"].Count() == 0)
                    {
                        CompanyId = Convert.ToInt32(Session["UserCompany"]);
                        DepartmentId = Convert.ToInt32(Session["UserDepartmentId"]);
                    }
                }
                #endregion

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectTaskTimeline(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate
                    , DurationSplit);
                jsonResponse = JObject.Parse(dataRequest);
                #endregion

                #region //Response
                if (jsonResponse["status"].ToString() == "errorForDA")
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "errorForDA",
                        msg = jsonResponse["msg"].ToString()
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        dataGroupDiagram = jsonResponse["dataGroupDiagram"]
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
        }
        #endregion

        #region //GetProject 取得專案資料
        [HttpPost]
        public void GetProject(int ProjectId = -1, int ProjectTaskId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProject(ProjectId, ProjectTaskId, "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        public void GetProjectTask(int TaskId = -1, int ProjectId = -1, string TaskType = "")
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectTask(TaskId, ProjectId, TaskType, "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

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

        #region //GetProjectReply 取得任務回報內容資料
        [HttpPost]
        public void GetProjectReply(int ProjectId = -1, int TaskId = -1, string ReplyType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectReply(ProjectId, TaskId, ReplyType
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

        #region //GetProjectTaskFile 取得使用者可存取檔案資料
        [HttpPost]
        public void GetProjectTaskFile(int ProjectTaskId = -1, int TaskId = -1, string FileName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectTaskFile(ProjectTaskId
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

        #region //GetProjectFile 取得使用者可存取檔案資料
        [HttpPost]
        public void GetProjectFile(int ProjectId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectFile(ProjectId
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

        #region //GetTrackFilterDefault
        [HttpPost]
        public void GetTrackFilterDefault()
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetTrackFilterDefault();
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddProjectTaskFile 任務檔案資料新增
        [HttpPost]
        public void AddProjectTaskFile(int TaskId = -1, int FileId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "update");

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
        #endregion

        #region //Update
        #region //UpdateTrackFilterDefault 使用者篩選預設值更新
        public void UpdateTrackFilterDefault(string FilterJsonData)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.UpdateTrackFilterDefault(FilterJsonData);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
                WebApiLoginCheck("ProjectTrackManagement", "update");

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

        #region //UpdateProjectTaskFileLevel 專案任務檔案資閱覽檔案等級更新
        [HttpPost]
        public void UpdateProjectTaskFileLevel(int ProjectTaskFileId = -1, int LevelId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "update");

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
                WebApiLoginCheck("ProjectTrackManagement", "update");

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
        #region //DeleteProjectTaskFile 專案任務檔案刪除
        [HttpPost]
        public void DeleteProjectTaskFile(int ProjectTaskFileId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "update");


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

        #region //Download
        #region //自定義儲存格
        public IXLStyle GenCellStyle(string status)
        {
            var xlstyle = XLWorkbook.DefaultStyle;
            xlstyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            xlstyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            xlstyle.Border.TopBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.BottomBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.LeftBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.RightBorder = XLBorderStyleValues.Thin;
            xlstyle.Border.TopBorderColor = XLColor.Black;
            xlstyle.Border.BottomBorderColor = XLColor.Black;
            xlstyle.Border.LeftBorderColor = XLColor.Black;
            xlstyle.Border.RightBorderColor = XLColor.Black;
            xlstyle.Font.FontName = "Microsoft JhengHei UI";   //微軟正黑體
            xlstyle.Font.FontSize = 12;
            xlstyle.Font.Bold = false;
            xlstyle.Protection.SetLocked(false);
            switch (status)
            {
                case "B":
                    xlstyle.Fill.BackgroundColor = XLColor.FromName("CadetBlue");
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
                case "I":
                    xlstyle.Fill.BackgroundColor = XLColor.Yellow;
                    break;
                case "C":
                    xlstyle.Fill.BackgroundColor = XLColor.Green;
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
                case "T":
                    xlstyle.Fill.BackgroundColor = XLColor.FromIndex(25);
                    xlstyle.Font.FontColor = XLColor.White;
                    break;
            }
            return xlstyle;
        }
        #endregion

        #region //GenProjectTaskToExcel
        public void GenProjectTaskToExcel(int ProjectId = -1, string ProjectUserId = "", string TaskUserId = "", string ProjectMaster = "", int CompanyId = -1, int DepartmentId = -1, string Project = "", string TaskName = ""
            , string ProjectStatus = "", string TaskStatus = "", string ReplyStatus = "", string TaskStartDate = "", string TaskEndDate = "", string TaskBaseDate = "")
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "excel");

                #region //Request authority
                systemSettingDA = new SystemSettingDA();
                dataRequest = systemSettingDA.GetAuthorityVerify("ProjectTrackManagement", "constrained-data");
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                if (jsonResponse["status"].ToString() == "success")
                {
                    if (jsonResponse["result"].Count() == 0)
                    {
                        CompanyId = Convert.ToInt32(Session["UserCompany"]);
                        DepartmentId = Convert.ToInt32(Session["UserDepartmentId"]);
                    }
                }
                #endregion

                #region //Request
                projectTrackDA = new ProjectTrackDA();
                dataRequest = projectTrackDA.GetProjectTaskTree(ProjectId, ProjectUserId, TaskUserId, ProjectMaster, CompanyId, DepartmentId, Project, TaskName
                    , ProjectStatus, TaskStatus, ReplyStatus, TaskStartDate, TaskEndDate, 0, "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                {
                    #region //樣式
                    #region //defaultStyle
                    var defaultStyle = XLWorkbook.DefaultStyle;
                    defaultStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    defaultStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    defaultStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                    defaultStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                    defaultStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                    defaultStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                    defaultStyle.Border.TopBorderColor = XLColor.Black;
                    defaultStyle.Border.BottomBorderColor = XLColor.Black;
                    defaultStyle.Border.LeftBorderColor = XLColor.Black;
                    defaultStyle.Border.RightBorderColor = XLColor.Black;
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
                    string excelFileName = "專案報表Excel檔" + TaskStartDate + "_" + TaskEndDate;
                    string excelsheetName = "專案任務";

                    List<ProjectList> data = JsonConvert.DeserializeObject<List<ProjectList>>(JObject.Parse(dataRequest)["data"].ToString());

                    //new string[] { "公司", "部門", "專案代碼", "專案名稱", "專案描述", "專案管理者", "工作天模式", "專案屬性", "專案狀態" }; //"父層任務"
                    string[] header = new string[] { "專案代碼", "專案名稱", "公司", "部門", "專案分類", "專案管理者", "專案成員", "工作天模式", "專案狀態", "專案描述",
                                "任務名稱", "任務描述", "任務狀態", "任務成員", "開始時間", "結束時間", "預估執行時間", "實際執行時間", "回報紀錄" };
                    string colIndex = "";
                    int rowIndex = 1;
                    #endregion

                    var personalTaskRequest = projectTrackDA.GetPersonalTaskFromProject(data.Select(s => Convert.ToInt32(s.ProjectId)).Distinct().ToList());
                    var personalTaskObj = new List<JObject>();
                    if (JObject.Parse(personalTaskRequest)["status"].ToString() == "success") {

                        personalTaskObj = JObject.Parse(personalTaskRequest)["data"].Children<JObject>().ToList();
                    }

                    #region //EXCEL
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(excelsheetName);
                        worksheet.RowHeight = 16;
                        worksheet.Style = defaultStyle;

                        #region //圖片
                        var imagePath = Server.MapPath("~/Content/images/logo.png");
                        var image = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 10)).Scale(0.07);
                        worksheet.Row(rowIndex).Height = 40;
                        worksheet.Range(rowIndex, 1, rowIndex, header.Length).Merge().Style = titleStyle;
                        rowIndex++;
                        #endregion

                        #region //HEADER
                        for (int i = 0; i < header.Length; i++)
                        {
                            colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            worksheet.Cell(colIndex).Value = header[i];
                            worksheet.Cell(colIndex).Style = headerStyle;
                        }
                        #endregion

                        bool first = true;
                        bool isBaseDate = DateTime.TryParse(TaskBaseDate, out DateTime taskBaseDate);
                        if (!isBaseDate) taskBaseDate = DateTime.Now;

                        #region //BODY
                        foreach (var proj in data)
                        {
                            var tasks = JArray.Parse(proj.ProjectTask); //project task array

                            if (first) { rowIndex++; first = false; }
                            worksheet.Cell(BaseHelper.MergeNumberToChar(1, rowIndex)).Value = proj.ProjectNo;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(2, rowIndex)).Value = proj.ProjectName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(3, rowIndex)).Value = proj.CompanyName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(4, rowIndex)).Value = proj.DepartmentName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(5, rowIndex)).Value = proj.ProjectTypeName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(6, rowIndex)).Value = proj.ProjectMaster;
                            //專案成員
                            worksheet.Cell(BaseHelper.MergeNumberToChar(7, rowIndex)).Value = string.IsNullOrWhiteSpace(proj.ProjectUser) ? "" :
                                    string.Join("\r\n", JObject.Parse(proj.ProjectUser)["data"]
                                        .Select(s => string.Format(@"{0}({1})", s["UserName"]?.ToString(), s["UserNo"]?.ToString())));
                            worksheet.Cell(BaseHelper.MergeNumberToChar(8, rowIndex)).Value = proj.WorkTimeStatusName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(9, rowIndex)).Value = proj.ProjectStatusName;
                            worksheet.Cell(BaseHelper.MergeNumberToChar(10, rowIndex)).Value = proj.ProjectDesc;
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = proj.ProjectAttributeName;

                            for (int i = 1; i < 11; i++)
                            {
                                worksheet.Range(rowIndex, i, rowIndex + tasks.Count() - 1, i).Merge().Style = defaultStyle;
                            }

                            //#region //HEADER
                            //for (int i = 0; i < header.Length; i++)
                            //{
                            //    colIndex = BaseHelper.MergeNumberToChar(i + 1, rowIndex);
                            //    worksheet.Cell(colIndex).Value = header[i];
                            //    worksheet.Cell(colIndex).Style = headerStyle;
                            //}
                            //#endregion

                            var projectTaskObj = tasks.Children<JObject>().OrderBy(o => o["TaskStart"]?.ToString()).ThenBy(t => t["TaskEnd"]?.ToString()).ToList();
                            var projectUserIds = JObject.Parse(proj.ProjectUser)["data"].Select(s=> s["UserId"]?.ToString());

                            if (personalTaskObj.Count > 0) projectTaskObj.AddRange(personalTaskObj
                                .Where(w => projectUserIds.Contains(w["UserId"]?.ToString()) 
                                    && Convert.ToDateTime(w["TaskStart"]?.ToString()).CompareTo(Convert.ToDateTime(proj.StartDate)) >= 0
                                    && Convert.ToDateTime(w["TaskStart"]?.ToString()).CompareTo(Convert.ToDateTime(proj.EndDate)) <= 0)
                                .OrderBy(o => o["TaskStart"]?.ToString())
                                .ThenBy(t => t["TaskEnd"]?.ToString()));

                            foreach (var task in projectTaskObj)//.OrderBy(o => o["ProjectNo"]?.ToString()).ThenBy(t => t["TaskStart"]?.ToString())
                            {
                                if (task["TaskType"]?.ToString() == "Personal") worksheet.Range(rowIndex, 1, rowIndex, 10).Merge().SetValue("個人任務");

                                DateTime.TryParse(task["TaskStart"]?.ToString(), out DateTime taskStart);
                                DateTime.TryParse(task["TaskEnd"]?.ToString(), out DateTime taskEnd);
                                double actualDuration = task["TaskStatus"]?.ToString() == "I" ? 
                                    taskBaseDate.Subtract(taskStart).TotalMinutes : 
                                    (task["TaskStatus"]?.ToString() == "C" ? Convert.ToDouble(task["TaskDuration"]?.ToString()) : -1);

                                worksheet.Cell(BaseHelper.MergeNumberToChar(11, rowIndex)).SetValue(task["TaskName"]?.ToString()); //任務名稱
                                worksheet.Cell(BaseHelper.MergeNumberToChar(12, rowIndex)).SetValue(task["TaskDesc"]?.ToString()); //任務描述
                                worksheet.Cell(BaseHelper.MergeNumberToChar(13, rowIndex)).SetValue(task["TaskStatusName"]?.ToString()).Style = GenCellStyle(task["TaskStatus"]?.ToString()); //任務狀態
                                //任務成員
                                worksheet.Cell(BaseHelper.MergeNumberToChar(14, rowIndex))
                                    .SetValue(task["TaskUser"]?.ToString() == "" ? "" :
                                        string.Join("\r\n", JObject.Parse(task["TaskUser"]?.ToString())["data"]
                                            .Select(s => string.Format(@"{0}({1})",s["UserName"]?.ToString(),s["UserNo"]?.ToString()))));
                                worksheet.Cell(BaseHelper.MergeNumberToChar(15, rowIndex)).SetValue(task["TaskStart"]?.ToString()); //開始時間
                                worksheet.Cell(BaseHelper.MergeNumberToChar(16, rowIndex)).SetValue(task["TaskEnd"]?.ToString()); //結束時間
                                worksheet.Cell(BaseHelper.MergeNumberToChar(17, rowIndex)).SetValue(GetDateTimeInfo(task["TaskDuration"]?.ToString())); //預估執行時間
                                worksheet.Cell(BaseHelper.MergeNumberToChar(18, rowIndex)).SetValue(GetDateTimeInfo(actualDuration.ToString())); //實際執行時間
                                //人:時間:狀態:延宕時間:內容
                                worksheet.Cell(BaseHelper.MergeNumberToChar(19, rowIndex)).Value = task["TaskReply"]?.ToString() == "" ? "" :
                                        string.Join("\r\n", JObject.Parse(task["TaskReply"]?.ToString())["data"]
                                            .Select(s => string.Format(@"{0}{1}{2}{3}{4}",
                                            s["UserName"]?.ToString(),
                                            s["ReplyDate"]?.ToString(),
                                            s["StatusName"]?.ToString(),
                                            GetDateTimeInfo(s["DeferredDuration"]?.ToString()),
                                            s["ReplyContent"]?.ToString())));
                                //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = task["TaskRouteName"]?.ToString();

                                rowIndex++;
                            }

                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = item.SoDetailRemark.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // 設定水平對齊為左對齊
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = item.RegularItemValue.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = item.FreebieItemValue.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = item.SpareItemValue.ToString();
                            //worksheet.Cell(BaseHelper.MergeNumberToChar(, rowIndex)).Value = " ";
                        }
                        #endregion

                        #region //設定
                        #region //篩選
                        worksheet.RangeUsed().SetAutoFilter(); // set filter
                        #endregion

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
                        worksheet.SheetView.FreezeRows(2);
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

                        #region //設定刻號欄寬
                        //worksheet.Column(11).Width = 50;
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

        #region //DownloadFile
        public ActionResult DownloadFile(int FileId, int ProjectId = -1)
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "download");

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
        #region //SendMessage MAMO訊息推播-New
        [HttpPost]
        [Route("api/SendMessage")]
        public void SendMessage(int TaskId = -1, string TaskType = "", string Content = "")
        {
            try
            {
                WebApiLoginCheck("ProjectTrackManagement", "read,constrained-data");

                #region //Request
                taskDA = new TaskDA();
                dataRequest = projectTrackDA.SendMessage(TaskId, TaskType, Content);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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