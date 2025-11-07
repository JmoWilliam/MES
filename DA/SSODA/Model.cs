using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSODA
{
    public class SettingDiagramShape
    {
        public string id { get; set; }
        public string flowName { get; set; }
        public string flowImage { get; set; }
        public string flowUser { get; set; }
        public int x { get; set; }
        public int y { get; set; }
        public string flowStatus { get; set; }
    }

    public class SourceFlowSetting
    {
        public int OriginalSettingId { get; set; }
        public int FlowId { get; set; }
        public int Xaxis { get; set; }
        public int Yaxis { get; set; }
        public int NewSettingId { get; set; }
    }

    public class SourceFlowLink
    {
        public int OriginalSourceSettingId { get; set; }
        public int OriginalTargetSettingId { get; set; }
        public int NewSourceSettingId { get; set; }
        public int NewTargetSettingId { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
    }

    public class SourceFlowUser
    {
        public int OriginalSettingId { get; set; }
        public int? RoleId { get; set; }
        public int? UserId { get; set; }
        public string MailAdviceStatus { get; set; }
        public string PushAdviceStatus { get; set; }
        public string WorkWeixinStatus { get; set; }
        public int NewSettingId { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
    }

    public class DemandFlow
    {
        public int SettingId { get; set; }
        public string FlowRoute { get; set; }
        public int FlowId { get; set; }
        public string FlowName { get; set; }
        public string FlowStatus { get; set; }
    }

    public class DemandFlowStatus
    {
        public int SettingId { get; set; }
        public string FlowStatus { get; set; }
    }
}
