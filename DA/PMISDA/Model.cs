using Helpers;
using System;
using System.Collections.Generic;

namespace PMISDA
{
    #region //樣板任務欄位
    public class TemlpateTask
    {
        public int TemplateTaskId { get; set; }
        public int? ParentTaskId { get; set; }
        public int? TaskLevel { get; set; }
        public string TaskRoute { get; set; }
        public int? TaskSort { get; set; }
        public string TaskName { get; set; }
        public int? StartPoint { get; set; }
        public int? Duration { get; set; }
        public int? ReplyFrequency { get; set; }
        public int SubTask { get; set; }
        public int? ProjectTaskId { get; set; }
        public int? ParentProjectTaskId { get; set; }
        public DateTime? StartDate { get; set; }
    }
    #endregion

    #region //樣板任務連結
    public class TemlpateTaskLink
    {
        public int TemplateTaskLinkId { get; set; }
        public int SourceTaskId { get; set; }
        public int TargetTaskId { get; set; }
        public string LinkType { get; set; }
    }
    #endregion

    #region //樣板任務成員
    public class TemplateTaskUser
    {
        public int TemplateTaskUserId { get; set; }
        public int TemplateTaskId { get; set; }
        public int UserId { get; set; }
        public int LevelId { get; set; }
        public int AgentSort { get; set; }
        public string Authority { get; set; }
    }
    #endregion

    #region //專案任務欄位
    public class ProjectTask
    {
        public int Index { get; set; }
        public string ProjectId { get; set; }
        public string ProjectStatus { get; set; }
        public float ProjectTaskId { get; set; }
        public string TaskRoute { get; set; }
        public int? TaskSort { get; set; }
        public string TaskName { get; set; }
        public string TaskStatus { get; set; }
        public string TaskStatusName { get; set; }
        public string TaskDesc { get; set; }
        public int? ParentTaskId { get; set; }
        public int? TaskLevel { get; set; }
        public int? ChannelId { get; set; }
        public DateTime? PlannedStart { get; set; }
        public DateTime? PlannedEnd { get; set; }
        public int? PlannedDuration { get; set; }
        public DateTime? EstimateStart { get; set; }
        public DateTime? EstimateEnd { get; set; }
        public int? EstimateDuration { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualEnd { get; set; }
        public int? ActualDuration { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public int? Duration { get; set; }
        public int? ReplyFrequency { get; set; }
        public bool RequiredFile { get; set; }
        public int SubUser { get; set; }
        public int SubTask { get; set; }
        public int UserTask { get; set; }
        public int DisplayTask { get; set; } = 0;
        public int DeferredStatus { get; set; } = 0;
    }
    #endregion

    public class DHXProjectTask : DHXTask
    {
        public string TaskRouteName { get; set; }
        public int? ReplyFrequency { get; set; }
        public string TaskUserData { get; set; }
        public string TaskUser { get; set; }
        public int SubTask { get; set; }
        public string TaskType { get; set; }
    }

    #region //專案任務連結
    public class ProjectTaskLink
    {
        public int? ProjectTaskLinkId { get; set; }
        public int SourceTaskId { get; set; }
        public int TargetTaskId { get; set; }
        public string LinkType { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //專案任務成員
    public class ProjectTaskUser
    {
        public int TaskUserId { get; set; }
        public int TaskId { get; set; }
        public int UserId { get; set; }
        public string UserNo { get; set; }
        public int LevelId { get; set; }
        public int AgentSort { get; set; }
        public string Authority { get; set; }
    }
    #endregion

    #region //專案清單欄位
    public class ProjectList
    {
        public string ProjectId { get; set; }
        public string ProjectNo { get; set; }
        public string ProjectName { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public int Duration { get; set; }
        public string ProjectPlannedStart { get; set; }
        public string ProjectPlannedEnd { get; set; }
        public int ProjectPlannedDuration { get; set; }
        public string ProjectDesc { get; set; }
        public string CompanyName { get; set; }
        public string CompanyNo { get; set; }
        public int LogoIcon { get; set; }
        public string DepartmentName { get; set; }
        public string DepartmentNo { get; set; }
        public string MasterUserNo { get; set; }
        public string MasterGender { get; set; }
        public string ProjectMaster { get; set; }
        public string ProjectUser { get; set; } // project user json string
        public int ProjectUserCount { get; set; }
        public string ProjectStatus { get; set; }
        public string ProjectStatusName { get; set; }
        public string ProjectType { get; set; }
        public string ProjectTypeName { get; set; }
        public string ProjectAttribute { get; set; }
        public string ProjectAttributeName { get; set; }
        public string WorkTimeStatus { get; set; }
        public string WorkTimeStatusName { get; set; }
        public string ProjectTask { get; set; } //task json string
        public double CompletionRate { get; set; } //完成進度
        public int TotalCount { get; set; }
        public DHXTree TaskTree { get; set; }
    }
    #endregion

    #region //任務清單欄位
    public class TaskList
    {
        public string ProjectId { get; set; }
        public string ProjectNo { get; set; }
        public string ProjectName { get; set; }
        public string ProjectTask { get; set; } //task json string
        public int TotalCount { get; set; }
        public DHXTree TaskTree { get; set; }
    }
    #endregion

    #region //時間軸圖節點
    public class DHXCircle
    {
        public string id { get; set; }
        public string type { get; set; }
        public string text { get; set; }
        public int width { get; set; }
        public int height { get; set; }
        public string fill { get; set; }
        public string stroke { get; set; }
        public string fontStyle { get; set; }
        public string fontColor { get; set; }
        public int x { get; set; }
        public int y { get; set; }
    }
    #endregion

    #region //時間軸顯示資訊
    public class TimeLineShapes : DHXCircle
    {
        public string TaskStatus { get; set; }
        public string IconPath { get; set; }
        public string Name { get; set; }
        public string TaskUser { get; set; }
        public string DateLine { get; set; }
    }

    public class GroupDiagram : ProjectList
    {
        public List<DHXCircle> dHXCircle { get; set; }
        public List<TimeLineShapes> dHXShapes { get; set; }
        public List<DHXDiagramLine> dHXDiagramLines { get; set; }
    }
    #endregion
}
