using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASDA
{
    public class User
    {
        public string UserType { get; set; }
        public string Account { get; set; }
        public string InnerCode { get; set; }
        public string DisplayName { get; set; }
        public Employee Employee { get; set; }
        public Customer Customer { get; set; }
        public Supplier Supplier { get; set; }
    }

    public class Employee
    {
        public int DepartmentId { get; set; }
        public string Gender { get; set; }
        public string Email { get; set; }
        public string Job { get; set; }
        public string JobType { get; set; }
        public string EmployeeStatus { get; set; }
    }

    public class Customer
    {
        public string Contact { get; set; }
        public string ContactPhone { get; set; }
        public string ContactAddress { get; set; }
    }

    public class Supplier
    {
        public string Contact { get; set; }
        public string ContactPhone { get; set; }
        public string ContactAddress { get; set; }
    }

    public class Interfacing
    {
        public int CompanyId { get; set; }
        public string InterfacingType { get; set; }

        public List<InterfacingConnection> Details { get; set; }
    }

    public class InterfacingConnection
    {
        public int ConnectionId { get; set; }
        public int InterfacingId { get; set; }
        public string ConnectionServer { get; set; }
        public string ConnectionAccount { get; set; }
        public string ConnectionPassword { get; set; }
        public string ConnectionDatabase { get; set; }
        public string Status { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
    }

    public class Operation
    {
        public int InterfacingId { get; set; }
        public string OperationNo { get; set; }
        public string OperationName { get; set; }
        public string Status { get; set; }
        public Import Import { get; set; }
        public Synchronize Synchronize { get; set; }
    }

    public class Import
    {
        public int SettingId { get; set; }
        public string Auto { get; set; }
        public int TimeInterval { get; set; }
        public string ExecuteMode { get; set; }
    }

    public class Synchronize
    {
        public int SettingId { get; set; }
        public string Auto { get; set; }
        public int TimeInterval { get; set; }
        public string ExecuteMode { get; set; }
    }

    #region //Mamo相關
    #region //建立團隊
    public class CreateTeam
    {
        public string Account { get; set; }
        public string TeamName { get; set; }
    }
    #endregion
    #endregion

    #region //共夾檔案上傳相關
    #region //資料夾
    public class Folder
    {
        public string FolderName { get; set; }
        public string FolderPath { get; set; }
    }
    #endregion

    #region //檔案路徑資料
    public class FilePathInfo
    {
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public byte[] FileByte { get; set; }
        public string FilePath { get; set; }
        public string FolderPath { get; set; }
        public string WhitelistPath { get; set; }
        public string ListId { get; set; }
    }
    #endregion
    #endregion
}
