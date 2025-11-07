using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpers
{
    public class DHXTree
    {
        public string status { get; set; }
        public string value { get; set; }
        public string id { get; set; }
        public bool opened { get; set; }
        public int level { get; set; }
        public List<DHXTree> items { get; set; }
    }

    public class DHXTask
    {
        public string id { get; set; }
        public string text { get; set; }
        public string start_date { get; set; }
        public string end_date { get; set; }
        public string duration { get; set; }
        public string progress { get; set; }
        public string parent { get; set; }
        public string type { get; set; }
        public int open { get; set; } = 1;
        public int order { get; set; }
        public int sub_user { get; set; }
        public string planned_start { get; set; }
        public string planned_end { get; set; }
        public string planned_duration { get; set; }
        public string task_status { get; set; }
        public string deferred_status { get; set; }
        public string task_status_name { get; set; }

        public bool allDay { get; set; }
        public string details { get; set; }
    }

    public class DHXLink
    {
        public string id { get; set; }
        public string type { get; set; }
        public string source { get; set; }
        public string target { get; set; }
    }

    public class DHXCollection
    {
        public List<DHXLink> links { get; set; }
    }

    public class DHXGantt
    {
        public List<DHXTask> data { get; set; }
        public DHXCollection collections { get; set; }
    }

    public class DHXDiagramLine
    {
        public string type { get; set; } = "line";
        public string id { get; set; }
        public string from { get; set; }
        public string fromSide { get; set; } = "right";
        public string to { get; set; }
        public string toSide { get; set; } = "left";
        public string connectType { get; set; } = "elbow";
        public string forwardArrow { get; set; } = "filled";
    }
}
