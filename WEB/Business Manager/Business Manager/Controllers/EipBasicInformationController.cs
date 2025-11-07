using BASDA;
using EIPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class EipBasicInformationController : WebController
    {
        private ScmBasicInformationDA scmBasicInformationDA = new ScmBasicInformationDA();
        private EipBasicInformationDA eipBasicInformationDA = new EipBasicInformationDA();

        #region //View
        public ActionResult DocNotificationReminder()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult DocNotificationUser()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult DocNotification()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetDocRoleInformation 取得角色單據資訊
        [HttpPost]
        public void GetDocRoleInformation(string Status="", string RoleName="", int RoleId=-1, string OrderBy="", int PageIndex=-1, int PageSize=-1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "read");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetDocRoleInformation(Status, RoleName, RoleId, OrderBy, PageIndex, PageSize);
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

        #region //GetDocUserInformation 取得使用者單據資訊
        [HttpPost]
        public void GetDocUserInformation(string UserName="", string RoleName = "", int RoleId = -1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "read");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetDocUserInformation(UserName, RoleName, RoleId, OrderBy, PageIndex, PageSize);
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

        #region //GetDocInformation 取得單據權限資訊
        [HttpPost]
        public void GetDocInformation(string DocType="", string RoleName = "", int RoleId = -1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "read");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetDocInformation(DocType, RoleName, RoleId, OrderBy, PageIndex, PageSize);
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

        #region //GetUser_S 取得人員在職資訊
        [HttpPost]
        public void GetUser_S(string DepartmentName="", int DepartmentId=-1, string UserName="", int UserId=-1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetUser_S(DepartmentName, DepartmentId, UserName, UserId, OrderBy, PageIndex, PageSize);
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

        #region //GetDepartment_A 取得部門啟用資訊
        [HttpPost]
        public void GetDepartment_A(string UserName="", int UserId=-1, string DepartmentName = "", int DepartmentId = -1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetDepartment_A(UserName, UserId, DepartmentName, DepartmentId, OrderBy, PageIndex, PageSize);
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

        #region //GetRfqProductType 取得RFQ資訊
        [HttpPost]
        public void GetRfqProductType(string RfqProductTypeName="", int RfqProTypeId=-1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetRfqProductType(RfqProductTypeName, RfqProTypeId, OrderBy, PageIndex, PageSize);
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

        #region //GetDfmItemCategory 取得DFM資訊
        [HttpPost]
        public void GetDfmItemCategory(string DfmItemCategoryName="", int DfmItemCategoryId=-1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetDfmItemCategory(DfmItemCategoryName, DfmItemCategoryId, OrderBy, PageIndex, PageSize);
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

        #region //GetMemberIcon -- 取得會員圖像
        public virtual ActionResult GetMemberIcon(int fileId)
        {
            try
            {
                #region //Request
                dataRequest = eipBasicInformationDA.GetMemberIcon(fileId, "", "", "", "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    string fileName = jsonResponse["result"][0]["FileName"].ToString();
                    byte[] fileContent = (byte[])jsonResponse["result"][0]["FileContent"];
                    string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();

                    return File(fileContent, FileHelper.GetMime(fileExtension), fileName + fileExtension);
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                return new EmptyResult();
            }
        }
        #endregion
        #endregion

        #region //Add
        #region //AddDocRoleInformation 單據角色資料新增
        [HttpPost]
        public void AddDocRoleInformation(string RoleName="")
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "add");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.AddDocRoleInformation(RoleName);
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

        #region //AddDocUserInformation 單據使用者資料新增
        [HttpPost]
        public void AddDocUserInformation(int RoleId=-1, int UserId=-1, string UserName="", string RoleName = "")
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "add");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.AddDocUserInformation(RoleId, UserId, UserName, RoleName);
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

        #region //AddDocInformation 單據權限資料新增
        [HttpPost]
        public void AddDocInformation(string SubProdType="", int SubProdId=-1, int RoleId = -1, string DocType = "", string RoleName = "")
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "add");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.AddDocInformation(SubProdType, SubProdId, RoleId, DocType, RoleName);
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
        #region //UpdateDocRoleInformation 單據角色資料更新
        [HttpPost]
        public void UpdateDocRoleInformation(int RoleId = -1, string RoleName = "")
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "update");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateDocRoleInformation(RoleId, RoleName);
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

        #region //UpdateDocRoleStatus 單據角色狀態更新
        [HttpPost]
        public void UpdateDocRoleStatus(int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "status-switch");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateDocRoleStatus(RoleId);
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
        #region //DeleteDocRoleInformation 單據角色刪除
        [HttpPost]
        public void DeleteDocRoleInformation(int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "delete");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.DeleteDocRoleInformation(RoleId);
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

        #region //DeleteDocUserInformation 單據人員刪除
        [HttpPost]
        public void DeleteDocUserInformation(int UserId = -1, int RoleId = -1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "delete");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.DeleteDocUserInformation(UserId, RoleId);
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

        #region //DeleteDocInformation 單據權限刪除
        [HttpPost]
        public void DeleteDocInformation(int DocNotifyId = -1)
        {
            try
            {
                WebApiLoginCheck("DocNotificationReminder", "delete");

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.DeleteDocInformation(DocNotifyId);
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

        #region //Api 電商
        #region //WebApiLoginCheckFromMember -- 取得登入狀態(客戶端) -- Chia Yuan --2023.07.26
        [NonAction]
        public void WebApiLoginCheckFromMember(int MemberType, string Account, string KeyText)
        {
            if (MemberType <= 0) throw new SystemException("【使用者類型】資料錯誤!");
            if (MemberType == 1) //客戶身分 23.08.21改採不認證直接詢價並自動登入，下次讓客戶更改密碼登入
            {
                #region //Request
                dataRequest = eipBasicInformationDA.GetLoginByKey(Account, KeyText);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "errorForDA")
                {
                    throw new SystemException(JObject.Parse(dataRequest)["msg"].ToString());
                }
            }
            if (MemberType == 2) //業務身分
            {
                #region //Request
                dataRequest = basicInformationDA.GetLoginByKey(Account, KeyText);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() == "errorForDA")
                {
                    throw new SystemException(JObject.Parse(dataRequest)["msg"].ToString());
                }
            }
        }
        #endregion

        #region //Get
        #region //GetMember -- 取得使用者資料 -- Chia Yuan -- 2023.07.27
        [HttpPost]
        [Route("api/EIP/GetMember")]
        public void GetMember(int MemberId = -1, string MemberName = "", string Status = ""
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetMember(MemberId, MemberName, Status
                    , MemberType, Account, KeyText
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

        #region //GetCustomer 取得客戶資料 --Chia Yuan -- 2023.08.01
        [HttpPost]
        [Route("api/EIP/GetCustomer")]
        public void GetCustomer(int CustomerId = -1, string CustomerNo = "", string CustomerName = "", string Status = ""
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();

                dataRequest = eipBasicInformationDA.GetCustomer(CustomerId, CustomerNo, CustomerName, Status, Account, KeyText
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

        #region //GetRfqProductClass -- 取得RFQ產品類型 -- Chia Yuan 2023.07.20
        [HttpPost]
        [Route("api/EIP/GetRfqProductClass")]
        public void GetRfqProductClass(int RfqProClassId = -1, string RfqProductClassName = "", string Status = ""
            , string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetRfqProductClass(RfqProClassId, RfqProductClassName, Status
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

        #region //GetRfqProductType 取得RFQ產品類別(客戶端) -- Chia Yuan -- 2023.07.11
        [HttpPost]
        [Route("api/EIP/GetRfqProductType")]
        public void GetRfqProductType(int RfqProTypeId = -1, int RfqProClassId = -1, string RfqProductTypeName = "", string Status = ""
            , string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetRfqProductType(RfqProTypeId, RfqProClassId, RfqProductTypeName, Status
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

        #region //GetRfqPackageType -- 取得RFQ包裝種類(客戶端 Cmb用) -- Chia Yuan 2023.07.12
        [HttpPost]
        [Route("api/EIP/GetRfqPackageType")]
        public void GetRfqPackageType(int RfqPkTypeId = -1, int RfqProClassId = -1, string PackagingMethod = "", string Status = "", string SustSupplyStatus = ""
            , string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetRfqPackageType(RfqPkTypeId, RfqProClassId, PackagingMethod, Status, SustSupplyStatus
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

        #region //GetProductUse -- 取得RFQ產品用途(客戶端 Cmb用) -- Chia Yuan -- 2023.7.11
        [HttpPost]
        [Route("api/EIP/GetProductUse")]
        public void GetProductUse(int ProductUseId = -1, string ProductUseNo = "", string ProductUseName = "", string Status = ""
            , string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = scmBasicInformationDA.GetProductUse(ProductUseId, ProductUseNo, ProductUseName, Status
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

        #region //GetRfqProductClassForList -- 取得RFQ產品類型(客戶端 清單用) -- Chia Yuan -- 2023.07.11
        [HttpPost]
        [Route("api/EIP/GetRfqProductClassForList")]
        public void GetRfqProductForList(int RfqProClassId = -1, string Status = ""
            , string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetRfqProductForList(RfqProClassId, Status
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

        #region //GetFileFromMember -- 取得檔案(客戶端) --Chia Yuan -- 2023.07.26
        [Route("api/EIP/GetFileFromMember")]
        public virtual ActionResult GetFileFromMember(int fileId)
        {
            try
            {
                #region //Request
                dataRequest = systemSettingDA.GetFile(fileId, -1, "N", "", "", "", -1, -1);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    string fileName = jsonResponse["result"][0]["FileName"].ToString();
                    byte[] fileContent = (byte[])jsonResponse["result"][0]["FileContent"];
                    string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();

                    return File(fileContent, FileHelper.GetMime(fileExtension), fileName + fileExtension);
                }
                else
                {
                    throw new SystemException(jsonResponse["msg"].ToString());
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                return new EmptyResult();
            }
        }
        #endregion

        #region //PreviewFromMember -- 檔案預覽(客戶端) --Chia Yuan -- 2023.07.26
        [HttpPost]
        [Route("api/EIP/PreviewFromMember")]
        public void PreviewFromMember(string previewFile)
        {
            try
            {
                string[] filePath = previewFile.Split(',');
                List<string> initialPreview = new List<string>();
                List<FileConfig> fileConfigs = new List<FileConfig>();

                if (filePath.Length > 0 && previewFile != "-1")
                {
                    for (int i = 0; i < filePath.Length; i++)
                    {
                        int fileId = Convert.ToInt32(filePath[i]);

                        #region //Request
                        dataRequest = systemSettingDA.GetFile(fileId, -1, "", "", "", "", -1, -1);
                        #endregion

                        #region //Response
                        jsonResponse = BaseHelper.DAResponse(dataRequest);
                        #endregion

                        if (jsonResponse["status"].ToString() == "success")
                        {
                            string fileName = jsonResponse["result"][0]["FileName"].ToString();
                            byte[] fileContent = (byte[])jsonResponse["result"][0]["FileContent"];
                            string fileExtension = jsonResponse["result"][0]["FileExtension"].ToString();
                            int fileSize = Convert.ToInt32(jsonResponse["result"][0]["FileSize"]);

                            string type = "";
                            bool filePreview = false;

                            FileHelper.FileType(ref type, ref filePreview, fileExtension);

                            string fileGuid = Guid.NewGuid().ToString();
                            Session[fileGuid] = fileContent;

                            string downloadPath = string.Format("/Web/Download?fileGuid={0}&fileName={1}&fileExtension={2}", fileGuid, fileName, fileExtension);

                            initialPreview.Add(downloadPath);

                            FileConfig fileConfig = new FileConfig
                            {
                                key = fileId.ToString(),
                                caption = fileName + fileExtension,
                                type = type,
                                filetype = FileHelper.GetMime(fileExtension),
                                size = fileSize,
                                previewAsData = filePreview,
                                url = "/Web/Delete",
                                downloadUrl = downloadPath
                            };

                            fileConfigs.Add(fileConfig);

                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "OK",
                                initialPreview,
                                initialPreviewConfig = fileConfigs,
                                initialPreviewAsData = true
                            });
                        }
                    }
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "error",
                        msg = "沒有可預覽的檔案!"
                    });
                }
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetRequestForQuotation 取得詢價清單列表(客戶端 RFQ)資訊 -- Chia Yuan -- 2023.07.11
        [HttpPost]
        [Route("api/EIP/GetRequestForQuotation")]
        public void GetRequestForQuotation(int RfqId = -1, string RfqNo = "", string MemberName = "", string AssemblyName = "", int ProductUseId = -1, int SalesId = -1, int RfqProTypeId = -1, string Status = ""
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetRequestForQuotation(RfqId, RfqNo, MemberName, AssemblyName, ProductUseId, SalesId, RfqProTypeId, Status
                    , MemberType, Account, KeyText
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

        #region //GetRfqDetail 取得RFQ單身資訊(客戶端) -- Chia Yuan -- 2023.07.11
        [HttpPost]
        [Route("api/EIP/GetRfqDetail")]
        public void GetRfqDetail(int RfqId = -1, int RfqDetailId = -1
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetRfqDetail(RfqId, RfqDetailId
                    , MemberType, Account, KeyText
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

        #region //GetSales 取得RFQ負責業務(客戶端 Cmb用) -- Chia Yuan -- 2023.07.11
        [HttpPost]
        [Route("api/EIP/GetSales")]
        public void GetSales(int UserId = -1, string UserName = ""
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetSales(UserId, UserName, OrderBy, PageIndex, PageSize);
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

        #region //GetStatus 取得狀態資料(客戶端 Chk用) -- Chia Yuan -- 2023.07.12
        [HttpPost]
        [Route("api/EIP/GetStatus")]
        public void GetStatus(string StatusSchema = "", string StatusNo = "", string StatusName = ""
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetStatus(StatusSchema, StatusNo, StatusName
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

        #region //GetType 取得類別資料(客戶端 Cmb用) -- Chia Yuan -- 2023.07.13
        [HttpPost]
        [Route("api/EIP/GetType")]
        public void GetType(string TypeSchema = "", string TypeNo = "", string TypeName = ""
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                basicInformationDA = new BasicInformationDA();
                dataRequest = basicInformationDA.GetType(TypeSchema, TypeNo, TypeName
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

        #region //GetRfqQtyLevel -- 取得RFQ報價方案(客戶端 Cmb用) -- Chia Yuan 2023.7.12
        [HttpPost]
        [Route("api/EIP/GetRfqQtyLevel")]
        public void GetRfqQtyLevel(int RfqQtyLevelId = -1
            , int MemberType = -1, string Account = "", string KeyText = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetRfqQtyLevel(RfqQtyLevelId
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
        #region //AddRequestForQuotation RFQ單頭新增 -- Chia Yuan 2023.7.17
        [HttpPost]
        [Route("api/EIP/AddRequestForQuotation")]
        public void AddRequestForQuotation(string AssemblyName = "", int ProductUseId = -1, string CustomerName = "", int CustomerId = -1, string Status = ""
            , string ContactName = "", string ContactPhone = ""
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //自動註冊

                #region //Request
                //if (MemberType == 1)
                //{
                //    eipBasicInformationDA = new EipBasicInformationDA();
                //    dataRequest = eipBasicInformationDA.RegisterUser(Account, Account, ContactName, CustomerName, ContactName, ContactPhone);
                //}
                #endregion

                #region //Response
                //jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #endregion

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.AddRequestForQuotation(AssemblyName, ProductUseId, CustomerName, CustomerId, Status, MemberType, Account, KeyText);
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

        #region //AddRfqDetail RFQ單身新增 -- Chia Yuan 2023.7.17
        [HttpPost]
        [Route("api/EIP/AddRfqDetail")]
        public void AddRfqDetail(int RfqId = -1, int RfqProTypeId = -1, string MtlName = "", int CustProdDigram = -1, int PrototypeQty = -1
            , string ProdLifeCycleStart = "", string ProdLifeCycleEnd = "", int LifeCycleQty = -1, string DemandDate = "", string CoatingFlag = "", int AdditionalFile = -1, string Description = ""
            , int RfqPkTypeId = -1, string RfqLineSolution = "", string Currency = "", string PortOfDelivery = "", string savePath = ""
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                List<FileModel> Files = FileHelper.FileSave(Request.Files);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.AddRfqDetail(RfqId, RfqProTypeId, MtlName, PrototypeQty, ProdLifeCycleStart, ProdLifeCycleEnd, LifeCycleQty, DemandDate, CoatingFlag, AdditionalFile, Description
                    , RfqPkTypeId, RfqLineSolution, Currency, PortOfDelivery
                    , Files, BaseHelper.ClientIP(), savePath
                    , MemberType, Account, KeyText);
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

        #region //AddRequestForQuotationFromJson --RFQ單頭一次新增 --Chia Yuan 2023.7.7
        [HttpPost]
        [Route("api/EIP/AddRequestForQuotationFromJson")]
        public void AddRequestForQuotationFromJson(string RfqEstimate = ""
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.AddRequestForQuotationFromJson(RfqEstimate, MemberType, Account, KeyText);
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
        #region //UpdateRequestForQuotation -- RFQ單頭更新 --Chia Yuan -- 2023.07.12
        [HttpPost]
        [Route("api/EIP/UpdateRequestForQuotation")]

        public void UpdateRequestForQuotation(int RfqId = -1, string AssemblyName = "", int ProductUseId = -1, string CustomerName = "", int CustomerId = -1, string Status = ""
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateRequestForQuotation(RfqId, AssemblyName, ProductUseId, CustomerName, CustomerId, Status, MemberType, Account, KeyText);
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

        #region //UpdateRfqDetail -- RFQ單身更新 -- Chia Yuan -- 2023.07.12
        [HttpPost]
        [Route("api/EIP/UpdateRfqDetail")]

        public void UpdateRfqDetail(int RfqDetailId = -1, int RfqProTypeId = -1, string MtlName = "", int CustProdDigram = -1, int PrototypeQty = -1
            , string ProdLifeCycleStart = "", string ProdLifeCycleEnd = "", int LifeCycleQty = -1, string DemandDate = "", string CoatingFlag = "", int AdditionalFile = -1, string Description = ""
            , int RfqPkTypeId = -1, string RfqLineSolution = "", string Currency = "", string PortOfDelivery = "", string savePath = "" //, HttpPostedFileBase[] fileCustProdDigram = null
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                List<FileModel> Files = FileHelper.FileSave(Request.Files);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateRfqDetail(RfqDetailId, RfqProTypeId, MtlName, PrototypeQty
                    , ProdLifeCycleStart, ProdLifeCycleEnd, LifeCycleQty, DemandDate, CoatingFlag, AdditionalFile, Description
                    , RfqPkTypeId, RfqLineSolution, Currency, PortOfDelivery
                    , Files, BaseHelper.ClientIP(), savePath
                    , MemberType, Account, KeyText);
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

        #region  //UpdateRequestForQuotationStatus -- RFQ單頭狀態更新 -- Chia Yuan -- 2023.07.21
        [HttpPost]
        [Route("api/EIP/UpdateRequestForQuotationStatus")]

        public void UpdateRequestForQuotationStatus(int RfqId = -1, string Status = ""
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateRequestForQuotationStatus(RfqId, Status
                    , MemberType, Account, KeyText);
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
        #region //DeleteRequestForQuotation -- RFQ單頭刪除 -- Chia Yuan -- 2023.07.22
        [HttpPost]
        [Route("api/EIP/DeleteRequestForQuotation")]
        public void DeleteRequestForQuotation(int RfqId = -1
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.DeleteRequestForQuotation(RfqId, MemberType, Account, KeyText);
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

        #region //DeleteRfqDetail -- RFQ單身刪除 -- Chia Yuan -- 2023.07.21
        [HttpPost]
        [Route("api/EIP/DeleteRfqDetail")]
        public void DeleteRfqDetail(int RfqDetailId = -1
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                scmBasicInformationDA = new ScmBasicInformationDA();
                dataRequest = eipBasicInformationDA.DeleteRfqDetail(RfqDetailId, MemberType, Account, KeyText);
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

        #region //DeleteFileFromMember -- 檔案刪除(客戶端) --Chia Yuan -- 2023.07.26
        [HttpPost]
        [Route("api/EIP/DeleteFileFromMember")]
        public void DeleteFileFromMember(string key, string Account, string KeyText)
        {
            try
            {
                #region //Request
                dataRequest = systemSettingDA.DeleteFile(Convert.ToInt32(key));
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    message = e.Message
                });

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion
        #endregion

        #region //Download
        #region //GetReturnRfqDetailDoc -- 取得報價單資料 --Chia Yuan -- 2023.07.31
        [HttpPost]
        [Route("api/EIP/GetReturnFile")]
        public void GetReturnFile(int FileId = -1
            , int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheckFromMember(MemberType, Account, KeyText);

                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.GetReturnRfqDetailDoc(FileId, Account, KeyText);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                if (jsonResponse["status"].ToString() == "success")
                {
                    byte[] wordFile;
                    string wordFileName = "", fileGuid = "", filePath = "", secondFilePath = "", user = ""
                        , password = BaseHelper.RandomCode(10);

                    var result = JObject.Parse(dataRequest);
                    if (result["data"].Count() <= 0) throw new SystemException("資料不完整!");
                    var data = result["data"][0];

                    //using (MemoryStream output = new MemoryStream())
                    //{
                    //    //doc.AddPasswordProtection(EditRestrictions.readOnly, password);
                    //    //doc.SaveAs(output, password);
                    //    wordFile = output.ToArray();
                    //}

                    wordFileName = string.Format("{0}", data["FileName"]);
                    fileGuid = Guid.NewGuid().ToString();
                    wordFile = (byte[])data["FileContent"];
                    Session.Add(fileGuid, wordFile);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        fileGuid,
                        fileName = wordFileName,
                        fileExtension = data["FileExtension"],
                        FileId
                        //file = wordFile,
                    });
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

        #region //DownloadFromMember -- 檔案下載 --Chia Yuan --2023.08.02
        [Route("api/EIP/DownloadFromMember")]
        public virtual ActionResult DownloadFromMember(int FileId, string fileGuid, string fileName, string fileExtension)
        {
            #region //Request
            eipBasicInformationDA = new EipBasicInformationDA();
            dataRequest = eipBasicInformationDA.GetFile(FileId);
            #endregion

            #region //Response
            jsonResponse = BaseHelper.DAResponse(dataRequest);
            #endregion

            //FileContent
            if (jsonResponse["status"].ToString() == "success")
            {
                //if (result["FileContent"] != null)

                var result = JObject.Parse(dataRequest);
                if (result["data"].Count() <= 0) throw new SystemException("資料不完整!");
                byte[] data = (byte[])result["data"][0]["FileContent"];

                return File(data, FileHelper.GetMime(fileExtension), fileName + fileExtension);
            }
            else
            {
                return new EmptyResult();
            }

            //if (Session[fileGuid] != null)
            //{
            //    byte[] data = Session[fileGuid] as byte[];
            //    return File(data, FileHelper.GetMime(fileExtension), fileName + fileExtension);
            //}
            //else
            //{
            //    return new EmptyResult();
            //}
        }
        #endregion
        #endregion
        #endregion

        #region //Api 客戶查詢系統

        #endregion
    }
}