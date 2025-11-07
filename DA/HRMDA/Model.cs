using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HRMDA
{
    #region //部門欄位
    public class Department
    {
        public int? DepartmentId { get; set; }
        public int? CompanyId { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public string SystemStatus { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //使用者欄位
    public class User
    {
        public int? UserId { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public string Gender { get; set; }
        public string Email { get; set; }
        public string MobilePhone { get; set; }        
        public string Password { get; set; }
        public string Job { get; set; }
        public string JobType { get; set; }
        public string UserStatus { get; set; }
        public string SystemStatus { get; set; }
        public string PasswordStatus { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //新版使用者欄位
    public class UserInfo
    {
        public int? UserId { get; set; }
        public string Account { get; set; }
        public string Password { get; set; }
        public string InnerCode { get; set; }
        public string DisplayName { get; set; }
        public string PasswordStatus { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //使用者內部員工設定檔
    public class UserEmployee
    {
        public int? UserId { get; set; }
        public string UserNo { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public string Gender { get; set; }
        public string Email { get; set; }
        public string Job { get; set; }
        public string JobType { get; set; }
        public string EmployeeStatus { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion
}
