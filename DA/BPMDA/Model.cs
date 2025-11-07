using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPMDA
{
    #region //BM userList
    public class UserList
    {
        public int UserId { get; set; }
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public string Gender { get; set; }
        public string Email { get; set; }
        public string DepartmentName { get; set; }
        public string Job { get; set; }
        public string JobType { get; set; }
        public int CompanyId { get; set; }
        public string CompanyName { get; set; }
    }
    #endregion

    #region //MES userList
    public class MESUserList
    {
        public int USER_ID { get; set; }
        public string BPM_USERNO { get; set; }
        public string BPM_USERNAME { get; set; }
       
    }
    #endregion
}
