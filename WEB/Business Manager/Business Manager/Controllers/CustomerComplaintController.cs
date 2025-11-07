using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using QMSDA;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Font = iTextSharp.text.Font;

namespace Business_Manager.Controllers
{
    public class CustomerComplaintController : WebController
    {
        private CustomerComplaintDA customerComplaintDA = new CustomerComplaintDA();

        public ActionResult CustomerComplaintManagement()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult CustomerComplaintMenuCategory()
        {
            ViewLoginCheck();

            return View();
        }


        #region//Get

        #region//GetCCMenuCategory --取得客訴選單類別
        [HttpPost]
        public void GetCCMenuCategory(int McId = -1, string McType = "", string McNo = "", string McName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintMenuCategory", "read,constrained-data");

                #region //Request
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCCMenuCategory(McId, McType, McNo, McName, Status
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

        #region//GetMenuCategorySimple --取得客訴選單類別(下拉用)
        [HttpPost]
        public void GetMenuCategorySimple(int McId = -1, string McType = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintMenuCategory", "read,constrained-data");

                #region //Request
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetMenuCategorySimple(McId, McType);
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

        #region//GetCustomerComplaint --取得客訴單據
        [HttpPost]
        public void GetCustomerComplaint(int CcId = -1, string CcNo = "", string MtlItemNo = "", int CustomerId = -1, string CurrentStatus = "", string StartDate = "", string EndDate = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "read,constrained-data");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCustomerComplaint(CcId, CcNo, MtlItemNo, CustomerId, CurrentStatus, StartDate, EndDate, Status
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

        #region//GetCcPdfShowSetting --取得客訴單據PDF顯示設定
        [HttpPost]
        public void GetCcPdfShowSetting(int CcId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "read");

                #region //Request
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCcPdfShowSetting(CcId);
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

        #region//GetCcStageData --取得客訴單據各階段資料
        [HttpPost]
        public void GetCcStageData(int CcStageDataId = -1, int CcId = -1, string Stage = "", int McId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "read");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCcStageData(CcStageDataId, CcId, Stage, McId
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

        #region//GetCcStageDataFile --取得客訴單據各階段資料檔案
        [HttpPost]
        public void GetCcStageDataFile(int CcStageDataFileId = -1, int CcStageDataId = -1, int CcId = -1, string Stage = "", string WhereFrom = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-read");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCcStageDataFile(CcStageDataId, CcStageDataId, CcId, Stage, WhereFrom
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

        #region//GetCCMember --取得D1小組成員
        [HttpPost]
        public void GetCCMember(int CcId = -1, string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "read");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCCMember(CcId, OrderBy, PageIndex, PageSize);
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

        #region//ShowCustomerComplaint ----客訴單檢視 
        [HttpPost]
        public void ShowCustomerComplaint(int CcId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "read");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.ShowCustomerComplaint(CcId);
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

        #region//GetCCMainFile --取得客訴單頭檔案
        [HttpPost]
        public void GetCCMainFile(int CcId = -1, string WhereFrom = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-read");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetCCMainFile(CcId, WhereFrom
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

        #region//GetD8TableData --取得D8附表
        [HttpPost]
        public void GetD8TableData(int CcStageDataId = -1, int CCTrackingDataId = -1, int CcId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "read");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.GetD8TableData(CcStageDataId, CCTrackingDataId, CcId);
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

        #region//Add
        #region //AddCCMenuCategory --新增選單類別
        [HttpPost]
        public void AddCCMenuCategory(string McName = "", string McNo = "", string McType = "", string McDesc = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintMenuCategory", "add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddCCMenuCategory(McName, McNo, McType, McDesc);
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

        #region//AddACustomerComplaint --新增客訴單據資料
        [HttpPost]
        public void AddACustomerComplaint(int MtlItemId = -1, string FilingDate = "", string DocDate = "", int CustomerId = -1
            , string UserName = "", string UserEMail = "", string IssueDescription = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddACustomerComplaint(MtlItemId, FilingDate, DocDate, CustomerId
                    , UserName, UserEMail, IssueDescription);
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

        #region//AddCcStageData --新增客訴單據各階段資料
        [HttpPost]
        public void AddCcStageData(int CcId = -1, string Stage = "", int McId = -1, int DeptId = -1, int UserId = -1, string DescriptionTextarea = "", string ExpectedDate ="", string FinishDate = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddCcStageData(CcId, Stage, McId, DeptId, UserId, DescriptionTextarea, ExpectedDate, FinishDate);
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

        #region//AddCcStageDataFile --新增客訴單據各階段檔案資料
        [HttpPost]
        public void AddCcStageDataFile(int CcStageDataId = -1, string FileId = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddCcStageDataFile(CcStageDataId, FileId);
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

        #region //AddCCUser --新增客訴單成員
        [HttpPost]
        public void AddCCUser(int CcId = -1, string UserId = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddCCUser(CcId, UserId);
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

        #region //AddCCMainFile --新增單頭檔案
        [HttpPost]
        public void AddCCMainFile(int CcId = -1, string FileId = "", string FileType = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddCCMainFile(CcId, FileId, FileType);
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

        #region //AddD8TableData --新增D8附表
        [HttpPost]
        public void AddD8TableData(int CcStageDataId = -1, int Batch = -1, string ReportNo = "", string TrialResult = "", string TrackingDate = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "add");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.AddD8TableData(CcStageDataId, Batch, ReportNo, TrialResult, TrackingDate);
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

        #region//Update
        #region//UpdateCCMenuCategory //編輯選單類別
        [HttpPost]
        public void UpdateCCMenuCategory(int McId = -1, string McName = "", string McNo = "", string McType = "", string McDesc = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintMenuCategory", "update");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateCCMenuCategory(McId, McName, McNo, McType, McDesc);
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

        #region//UpdateCustomerComplaint --更新客訴單據各階段資料
        [HttpPost]
        public void UpdateCustomerComplaint(int CcId = -1, string CcNo = "", int MtlItemId = -1, string FilingDate = "", string DocDate = "", int CustomerId = -1
            , string UserName = "", string UserEMail = "", string IssueDescription = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "update");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateCustomerComplaint(CcId, CcNo, MtlItemId, FilingDate, DocDate, CustomerId
                    , UserName, UserEMail, IssueDescription);
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

        #region//UpdateChangeCustomer -- 更新客訴單據客戶代碼替換資料
        [HttpPost]
        public void UpdateChangeCustomer(int CcId = -1, int ChangeCustomerId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "change-customer");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateChangeCustomer(CcId, ChangeCustomerId);
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

        #region//UpdatePdfShowSetting --更新PDF顯示狀態設定
        [HttpPost]
        public void UpdatePdfShowSetting(int CcId = -1, string Stage = "", string Model ="")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "status-switch");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdatePdfShowSetting(CcId, Stage, Model);
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

        #region//UpdateStageDataSetting --更新階段維護設定
        [HttpPost]
        public void UpdateStageDataSetting(int CcId = -1, string Stage = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "status-switch");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateStageDataSetting(CcId, Stage);
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

        #region//UpdateCcStageData --更新客訴單據各階段資料
        [HttpPost]
        public void UpdateCcStageData(int CcStageDataId = -1, int McId = -1, int DeptId = -1, int UserId = -1, string DescriptionTextarea = "", string ExpectedDate = "", string FinishDate = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "update");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateCcStageData(CcStageDataId, McId, DeptId, UserId, DescriptionTextarea, ExpectedDate, FinishDate);
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

        #region //UpdateStageDataStatus --確認/反確認各階段資料 --GPai 20230918
        [HttpPost]
        public void UpdateStageDataStatus(int CcId = -1, string Stage = "", string StageStatus = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "update");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateStageDataStatus(CcId, Stage, StageStatus);
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

        #region //UpdateD8TableData --確認/反確認各階段資料 --GPai 20230918
        [HttpPost]
        public void UpdateD8TableData(int CCTrackingDataId = -1, int Batch =-1, string ReportNo = "", string TrialResult = "", string TrackingDate = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "reconfirm");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateD8TableData(CCTrackingDataId, Batch, ReportNo, TrialResult, TrackingDate);
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

        #region //UpdateCCMenuCategoryStatus //選單類別狀態
        [HttpPost]
        public void UpdateCCMenuCategoryStatus(int McId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintMenuCategory", "update");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.UpdateCCMenuCategoryStatus(McId);
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

        #region//Delete
        #region//DeleteCCMenuCategory //刪除選單類別
        [HttpPost]
        public void DeleteCCMenuCategory(int McId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintMenuCategory", "delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteCCMenuCategory(McId);
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

        #region//DeleteCustomerComplaint --刪除客訴單據
        [HttpPost]
        public void DeleteCustomerComplaint(int CcId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteCustomerComplaint(CcId);
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

        #region//DeleteCcStageData --刪除客訴單據各階段資料
        [HttpPost]
        public void DeleteCcStageData(int CcId = -1,int CcStageDataId = -1,string Stage = "")
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteCcStageData(CcId, CcStageDataId, Stage);
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

        #region//DeleteCcStageDataFile --刪除客訴單據各階段資料檔案
        [HttpPost]
        public void DeleteCcStageDataFile(int CcStageDataFileId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteCcStageDataFile(CcStageDataFileId);
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

        #region//DeleteCCUser --刪除客訴單成員
        [HttpPost]
        public void DeleteCCUser(int UserId = -1, int CcId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteCCUser(UserId, CcId);
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

        #region //DeleteCCMainFile --刪除單頭檔案
        [HttpPost]
        public void DeleteCCMainFile(int CcFileId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "file-delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteCCMainFile(CcFileId);
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

        #region //DeleteD8TableData --刪除D8附表
        [HttpPost]
        public void DeleteD8TableData(int CCTrackingDataId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "delete");

                #region //Request 
                customerComplaintDA = new CustomerComplaintDA();
                dataRequest = customerComplaintDA.DeleteD8TableData(CCTrackingDataId);
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

        #region //PDF

        #region //CustomerComplaintCardPdf 客訴單據卡
        public void CustomerComplaintCardPdf(int CcId = -1, int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("CustomerComplaintManagement", "print");

                #region //產生PDF
                WebClient wc = new WebClient();

                #region //Request - 客訴單據卡資料
                dataRequest = customerComplaintDA.CustomerComplaintCardPdf(CcId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                #region //Request - 各階段確認狀態
                JObject jsonResponseConfirmStatus = new JObject();
                string dataRequestConfirmStatus = customerComplaintDA.GetDataConfirmStatus(CcId);
                jsonResponseConfirmStatus = BaseHelper.DAResponse(dataRequestConfirmStatus);
                #endregion

                #region //Request - 客訴單據卡D8資料
                JObject jsonResponseD8 = new JObject();
                string dataRequestD8 = customerComplaintDA.GetD8TableData(-1, -1, CcId);
                jsonResponseD8 = BaseHelper.DAResponse(dataRequestD8);
                #endregion

                #region //Request - 客訴單據卡圖片資料
                JObject jsonResponseImg = new JObject();
                string dataRequestImg = customerComplaintDA.GetCcStageDataFilePDf(CcId);
                jsonResponseImg = BaseHelper.DAResponse(dataRequestImg);
                #endregion

                string D8McName = "";
                bool Appendix = false;
                string CcNo = "";
                string MtlItemNo = "";
                string MtlItemName = "";
                string MtlItemSpec = "";
                string CustomerName = "";
                string htmlText = "";
                string CompanyName = "";
                string CodeText = "";
                string IssueDescription = "";
                string CcTeamMembers = "";
                string DocDate = "";
                string FilingDate = "";
                string Content = "";
                string D1Content = "";
                string D2Content = "";
                string D3Content = "";
                string D4Content = "";
                string D5Content = "";
                string D6Content = "";
                string D7Content = "";
                string D8Content = "";
                string D8DescriptionTextarea = "";
                string ConfirmInfoD2 = "";
                string ConfirmInfoD3 = "";
                string ConfirmInfoD4 = "";
                string ConfirmInfoD5 = "";
                string ConfirmInfoD6 = "";

                int D0Row = 0;
                int D1Row = 0;
                int D2Row = 0;
                int D3Row = 0;
                int D4Row = 0;
                int D5Row = 0;
                int D6Row = 0;
                int D7Row = 0;
                int D8Row = 0;

                if (CompanyId == 2)
                {
                    CompanyName = "中揚光電股份有限公司";
                    CodeText = "P-Q009表01-02";
                }
                else if (CompanyId == 3)
                {
                    CompanyName = "紘立光電股份有限公司";
                    CodeText = "Q-17-表01";
                }
                else if (CompanyId == 4)
                {
                    CompanyName = "晶彩光學有限公司";
                }
                else
                {
                    CompanyName = "中揚光電股份有限公司(錯誤)";
                }

                List<string> PDFShowSettingList = new List<string>();


                if (jsonResponse["status"].ToString() == "success")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    if (result.Count() <= 0) throw new SystemException("此單據無法進行列印作業!");
                    D8McName = result[0]["D8McName"].ToString();
                    if(D8McName == "報告回覆日期後，連續三批出貨無異常，予以結案")
                    {
                        Appendix = true;
                    }
                    
                    var resultConfirmStatus = JObject.Parse(dataRequestConfirmStatus)["data"];
                    dynamic[] dataConfirmStatus = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequestConfirmStatus)["data"].ToString());
                    string ConfirmStatusStr = "NNNNNNNN";
                    int ConfirmStatusStr1 = resultConfirmStatus.Count();
                    for (int i = 1; i <= resultConfirmStatus.Count(); i++)
                    {
                        string StageDB = resultConfirmStatus[i-1]["Stage"].ToString();
                        string ConfirmStatus = resultConfirmStatus[i-1]["ConfirmStatus"].ToString(); ;
                        int place = Convert.ToInt32(StageDB[StageDB.Length - 1].ToString());

                        ConfirmStatusStr = ConfirmStatusStr.Remove(place - 1, 1).Insert(place - 1, ConfirmStatus);


                        //ConfirmStatusStr = ConfirmStatusStr.Substring(0, place-1) + ConfirmStatus + ConfirmStatusStr.Substring(7- place-1);
                    }


                    #region //html

                    #region //PDF顯示設定
                    for (int i = 1; i <= 8; i++)
                    {
                        string ShowSettingData = "D" + i + "ShowSetting";
                        PDFShowSettingList.Add(result[0][ShowSettingData].ToString());
                    }
                    #endregion

                    #region //D1hmtl結構
                    if (PDFShowSettingList[0] == "A" && ConfirmStatusStr.Substring(1, 1) == "Y")
                    {
                        CcTeamMembers = result[0]["CcTeamMembers"].ToString();
                        //CcTeamMembers = "瞰著壯觀的風景山山川在山頂俯瞰著壯觀的風景山山川瞰著壯觀的風景山山川起伏山雲在山頂俯瞰著壯觀的風瞰著壯觀的風景山山川在山頂俯瞰著壯觀的風景山山川瞰著壯觀的風景山山川起伏山雲在山頂俯瞰著壯觀的風瞰著壯觀的風景山山川在山頂俯瞰著壯觀的風景山山川瞰著壯觀的風景山山川起伏山雲在山頂俯瞰著壯觀的風";
                        D1Content = @"
                            <tr>
                                <td colspan=""25"" rowspan=""1"" style=""height:48px;font-size:14px;"">
                                " + CcTeamMembers + @"
                                </td>
                            </tr>
                            ";
                        D1Row += CcTeamMembers.Length / 50 + (CcTeamMembers.Length % 50 != 0 ? 1 : 0);
                        D1Row = D1Row > 2 ? D1Row : 2;
                    }
                    else
                    {
                        D1Content = "";
                        D1Row = 0;
                    }
                    #endregion

                    #region //D2~D7hmtl結構
                    string[] stageDataArray = new string[7];

                    for (int i = 2; i <= 8; i++)
                    {
                        string propertyName = "StageDataD" + i;
                        dynamic stageData = result[0][propertyName];
                        stageDataArray[i - 2] = stageData != null ? stageData.ToString() : "";
                        string content = "";
                        int row = 0;
                        if (PDFShowSettingList[i-1] == "A" && ConfirmStatusStr.Substring(i - 1, 1) == "Y")
                        {
                            content = DealWithContent(stageData.ToString(), i, Appendix);
                            row = Convert.ToInt32(content.Split('♡')[0]);
                        }

                        switch (i)
                        {
                            case 2:
                                D2Row = row;
                                D2Content = content;
                                break;
                            case 3:
                                D3Row = row;
                                D3Content = content;
                                break;
                            case 4:
                                D4Row = row;
                                D4Content = content;
                                break;
                            case 5:
                                D5Row = row;
                                D5Content = content;
                                break;
                            case 6:
                                D6Row = row;
                                D6Content = content;
                                break;
                            case 7:
                                D7Row = row;
                                D7Content = content;
                                break;
                            case 8:
                                D8Row = row;
                                if (Appendix)
                                {
                                    D8DescriptionTextarea = content.Split('♡')[1];
                                }
                                else
                                {
                                    D8Content = content;
                                }
                                break;
                        }
                    }
                    #endregion

                    #region //D8hmtl結構
                    if (Appendix)
                    {
                        if (PDFShowSettingList[7] == "A" && ConfirmStatusStr.Substring(7, 1) == "Y")
                        {
                            var resultD8 = JObject.Parse(dataRequestD8)["data"];
                            dynamic[] dataD8 = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequestD8)["data"].ToString());

                            string TrackInfo = "";
                            int D8Count = 0;

                            #region //D8資料結構生成
                            foreach (var item in dataD8)
                            {
                                string[] CountArr = new string[] { "一", "二", "三" };
                                TrackInfo += @"
                                        <tr>
                                            <td style=""width:80px;border:0px;"">&nbsp;</td>
                                            <td>第" + CountArr[D8Count] + @"批</td>
                                            <td>" + item.TrackingDate + @"</td>
                                            <td>" + item.ReportNo + @"</td>
                                            <td>" + item.TrialResult + @"</td>
                                            <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        </tr>
                                    ";
                                D8Count++;
                            }
                            #endregion

                            #region //D8hmtl資料補足結構
                            if (D8Count == 0)
                            {
                                TrackInfo += @"
                                    <tr>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        <td>第一批</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        <td>第二批</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        <td>第三批</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                    </tr>
                                    ";
                            }
                            else if (D8Count == 1)
                            {
                                TrackInfo += @"
                                    <tr>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        <td>第二批</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        <td>第三批</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                    </tr>
                                    ";
                            }
                            else if (D8Count == 2)
                            {
                                TrackInfo += @"
                                    <tr>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                        <td>第三批</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                        <td style=""width:80px;border:0px;"">&nbsp;</td>
                                    </tr>
                                    ";
                            }
                            #endregion

                            #region //D8hmtl結構
                            D8Content = @"
                            <tr style='border-bottom:0px;'>
                                <td colspan=""25"" rowspan=""1"" style='border-bottom:0px;'>說明:[D8DescriptionTextarea]</td>
                            </tr>
                            <tr>
                                <td colspan=""25"" rowspan=""8"" style='height:165px;border-top:0px;border-bottom:0px;'>
                                    <table class=""minTable"" style=""width:100%;"">
                                        <tr>
                                            <td style=""width:10%;border:0px;"">&nbsp;</td>
                                            <td style=""width:15%;text-align:center;"">批數</td>
                                            <td style=""width:15%;text-align:center;"">追蹤日期</td>
                                            <td style=""width:30%;text-align:center;"">IQC/OQC報告編號</td>
                                            <td style=""width:20%;text-align:center;"">檢驗結果</td>
                                            <td style=""width:10%;border:0px;"">&nbsp;</td>
                                        </tr>
                                        [TrackInfo]
                                    </table>
                                </td>
                            </tr>
                            ";
                            D8Content = D8Content.Replace("[TrackInfo]", TrackInfo);
                            D8Content = D8Content.Replace("[D8DescriptionTextarea]", D8DescriptionTextarea);
                            #endregion
                        }
                        else
                        {
                            D8Content = "";
                            D8Row = 0;
                        }
                    }
                    #endregion

                    #region //各區域資料操作

                    #region //基本欄位取值
                    CcNo = result[0]["CcNo"].ToString();
                    MtlItemNo = result[0]["MtlItemNo"].ToString();
                    MtlItemName = result[0]["MtlItemName"].ToString();
                    CustomerName = result[0]["CustomerName"].ToString();
                    IssueDescription = result[0]["IssueDescription"].ToString();
                    MtlItemSpec = result[0]["MtlItemSpec"].ToString();
                    DocDate = result[0]["DocDate"].ToString();
                    FilingDate = result[0]["FilingDate"].ToString();
                    #endregion

                    htmlText = System.IO.File.ReadAllText(Server.MapPath("~/PdfTemplate/MES/CustomerComplaintCard.html"));
                    htmlText = htmlText.Replace("[CcNo]", CcNo);
                    htmlText = htmlText.Replace("[MtlItemNo]", MtlItemNo);
                    htmlText = htmlText.Replace("[MtlItemName]", MtlItemName);
                    htmlText = htmlText.Replace("[MtlItemSpec]", MtlItemSpec);
                    htmlText = htmlText.Replace("[DocDate]", DocDate);
                    htmlText = htmlText.Replace("[FilingDate]", FilingDate);
                    htmlText = htmlText.Replace("[IssueDescription]", IssueDescription);
                    htmlText = htmlText.Replace("[CustomerName]", CustomerName);

                    #region //各區域換頁計算
                    int rowTotal = 0;

                    D0Row += IssueDescription.Length / 50 + (IssueDescription.Length % 50 != 0 ? 1 : 0);
                    D0Row = D0Row > 5 ? D0Row : 5;
                    rowTotal += D0Row;
                    rowTotal += D1Row;
                    rowTotal += D2Row;
                    rowTotal += D3Row;
                    rowTotal += D4Row;
                    rowTotal += D5Row;
                    rowTotal += D6Row;
                    rowTotal += D7Row;
                    rowTotal += D8Row;

                    #region //D0~D4頁
                    int num1 = D0Row + D1Row + D2Row + D3Row + D4Row;
                    int num2 = D0Row + D1Row + D2Row + D3Row;
                    int num3 = D0Row + D1Row + D2Row;
                    int num4 = D0Row + D1Row;
                    int num5 = D0Row;

                    string pageChangeD1 = "auto";
                    string pageChangeD2 = "auto";
                    string pageChangeD3 = "auto";
                    string pageChangeD4 = "auto";
                    string pageChangeD6 = "auto";
                    string pageChangeD7 = "auto";
                    string pageChangeD8 = "auto";

                    if (num1 <= 31) //5
                    {
                        pageChangeD1 = pageChangeD2 = pageChangeD3 = pageChangeD4 = "auto";
                    }
                    else if (num2 <= 35) //4+1
                    {
                        pageChangeD1 = pageChangeD2 = pageChangeD3 = "auto";
                        pageChangeD4 = "always";
                    }
                    else if (num3 <= 38) //3+2
                    {
                        pageChangeD1 = pageChangeD2 = "auto";
                        pageChangeD3 = "always";
                        pageChangeD4 = (D3Row + D4Row) <= 53 ? "auto" : "always"; // 2;1+1
                    }
                    else if (num4 < 45) //2+3
                    {
                        pageChangeD1 = "auto";
                        pageChangeD2 = "always";

                        if ((D2Row + D3Row + D4Row) <= 51) //3
                        {
                            pageChangeD3 = pageChangeD4 = "auto";
                        }
                        else if ((D2Row + D3Row) <= 56) //2+1
                        {
                            pageChangeD3 = "auto";
                            pageChangeD4 = "always";
                        }
                        else if (D2Row > 50) //1+2
                        {
                            pageChangeD2 = "always";
                            pageChangeD3 = "always";
                            pageChangeD4 = (D3Row + D4Row) <= 53 ? "auto" : "always"; // 2;1+1
                        }
                    }
                    else if (D0Row >= 45) //1+4
                    {
                        pageChangeD1 = "always";

                        if ((D1Row + D2Row + D3Row + D4Row) <= 45) //4
                        {
                            pageChangeD2 = pageChangeD3 = pageChangeD4 = "auto";
                        }
                        else if ((D1Row + D2Row + D3Row) <= 51) //3+1
                        {
                            pageChangeD2 = pageChangeD3 = "auto";
                            pageChangeD4 = "always";
                        }
                        else if ((D1Row + D2Row) <= 53) //2+2
                        {
                            pageChangeD2 = "auto";
                            pageChangeD3 = "always";
                            pageChangeD4 = (D3Row + D4Row) <= 53 ? "auto" : "always"; // 2;1+1
                        }
                        else if(D2Row > 50)//1+1+2
                        {
                            pageChangeD2 = "always";
                            pageChangeD4 = (D3Row + D4Row) <= 53 ? "auto" : "always"; // 2;1+1

                            //2區塊 //1區塊 1區塊
                            //if ((D3Row + D4Row) <= 53) 
                            //{
                            //    pageChangeD4 = "auto";
                            //}
                            //else 
                            //{
                            //    pageChangeD4 = "always";

                            //}
                        }
                    }

                    htmlText = htmlText.Replace("[PageChangeD1]", pageChangeD1);
                    htmlText = htmlText.Replace("[PageChangeD2]", pageChangeD2);
                    htmlText = htmlText.Replace("[PageChangeD3]", pageChangeD3);
                    htmlText = htmlText.Replace("[PageChangeD4]", pageChangeD4);
                    #endregion

                    #region //D5~D8頁
                    int num6 = D5Row + D6Row + D7Row + D8Row;
                    int num7 = D5Row + D6Row + D7Row;
                    int num8 = D5Row + D6Row;
                    int num9 = D5Row;

                    if(num6 <= 20) //4
                    {
                        pageChangeD6 = pageChangeD7 = pageChangeD8 = "auto";
                    }
                    else if(num7 <= 42) //3+1
                    {
                        pageChangeD6 = pageChangeD7 = "auto";
                        pageChangeD8 = "always";
                    }
                    else if (num8 <= 51) //2+2
                    {
                        pageChangeD6 = "auto";
                        pageChangeD7 = "always";
                        pageChangeD8 = (D7Row + D8Row) <= 52 ? "auto" : "always"; // 2;1+1
                    }
                    else if (num9 >= 54) //1+3
                    {
                        pageChangeD6 = "always";
                        if((D6Row + D7Row + D8Row) <= 33)
                        {
                            pageChangeD7 = pageChangeD8 = "auto";
                        }
                        else if ((D6Row + D7Row) <= 56)
                        {
                            pageChangeD7 = "auto";
                            pageChangeD8 = "auto";
                        }
                        else if (D6Row >= 58)
                        {
                            pageChangeD7 = "always";
                            pageChangeD8 = (D7Row + D8Row) <= 52 ? "auto" : "always";
                        }
                    }
                    htmlText = htmlText.Replace("[PageChangeD6]", pageChangeD6);
                    htmlText = htmlText.Replace("[PageChangeD7]", pageChangeD7);
                    htmlText = htmlText.Replace("[PageChangeD8]", pageChangeD8);
                    #endregion

                    #endregion

                    #region //各區域資料帶入
                    int count = 1;
                    foreach (var item in PDFShowSettingList)
                    {
                        string ContentA = "[D" + count + "Content]";
                        string ContentB = "";

                        switch (count)
                        {
                            case 1:
                                ContentB = D1Content;
                                break;
                            case 2:
                                ContentB = D2Content;
                                break;
                            case 3:
                                ContentB = D3Content;
                                break;
                            case 4:
                                ContentB = D4Content;
                                break;
                            case 5:
                                ContentB = D5Content;
                                break;
                            case 6:
                                ContentB = D6Content;
                                break;
                            case 7:
                                ContentB = D7Content;
                                break;
                            case 8:
                                ContentB = D8Content;
                                break;
                        }
                        if (item == "A")
                        {

                            htmlText = htmlText.Replace(ContentA, ContentB);
                        }
                        else
                        {
                            htmlText = htmlText.Replace(ContentA, "");
                        }
                        count++;
                    }
                    #endregion
                    #endregion

                    #region //附件圖片區域
                    if (jsonResponseImg["status"].ToString() == "success")
                    {
                        if (jsonResponseImg["result"].ToString() != "[]")
                        {
                            string htmlDetail1 = "";

                            var resultImg = JObject.Parse(dataRequestImg)["data"];
                            dynamic[] data = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequestImg)["data"].ToString());

                            string fileName = "";
                            string fileContent = "";
                            string fileExtension = "";
                            int imgCount = 1;
                            foreach (var item in data)
                            {
                                fileName = item.FileName;
                                fileContent = item.FileContent;
                                fileExtension = item.FileExtension;
                                byte[] imageBytes = Convert.FromBase64String(fileContent);
                                MemoryStream memoryStream = new MemoryStream(imageBytes, 0, imageBytes.Length);
                                memoryStream.Write(imageBytes, 0, imageBytes.Length);
                                //二進位制轉成圖片儲存
                                System.Drawing.Image image = System.Drawing.Image.FromStream(memoryStream);
                                image.Save(Server.MapPath(string.Format("~/PdfTemplate/MES/CustomerComplaintCardImg/{0}.bmp", fileName)));


                                #region //附件html結構
                                string htmlCode = @"
                                            <div style=""page-break-after: always;"">
                                                <table class=""rtable"">
                                                    <tr>
                                                        <th colspan=""19"" rowspan=""2"" style=""width:76%;font-size:28px;border-bottom:0px;"">品質異常單</th>
                                                        <th colspan=""6"" style=""width:24%;text-align:center;font-size:14px;text-align:center;"">管理NO</th>
                                                    </tr>
                                                    <tr>
                                                        <td colspan=""6"" style=""font-size:14px;text-align:center;"">" + CcNo + @"</td>
                                                    </tr>
                                                    <tr>
                                                        <th colspan=""25"" rowspan=""1"" style=""width:100%;border-top:0px;border-bottom:0px;text-align:center;font-size:28px"">附件</th>
                                                    </tr>
                                                    <tr>
                                                        <td colspan=""25"" rowspan=""1"" style=""width:100%;height:60px;border-top:0px;border-bottom:0px;font-size:18px"">附件" + imgCount +"-"+ fileName+ @"</td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan=""25"" style=""width:100%;height:800px;border-top:0px;"">
                                                            <img style='width:680px;' src='" + (Server.MapPath("~/PdfTemplate/MES/CustomerComplaintCardImg/" + fileName + ".bmp").ToString()) + @"' alt=''/>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </div>
                                            ";
                                #endregion

                                htmlDetail1 += htmlCode;
                                imgCount++;
                            }

                            htmlText = htmlText.Replace("[imgArea]", htmlDetail1 != "" ? htmlDetail1 : "");

                        }
                        else
                        {
                            htmlText = htmlText.Replace("[imgArea]", "");
                        }
                    }
                    else
                    {
                        throw new SystemException(jsonResponse["msg"].ToString());
                    }

                    #endregion
                    #endregion
                }

                #region //匯出
                byte[] pdfFile;

                using (MemoryStream output = new MemoryStream()) //要把PDF寫到哪個串流
                {
                    byte[] data = Encoding.UTF8.GetBytes(htmlText); //字串轉成byte[]

                    using (MemoryStream input = new MemoryStream(data))
                    {
                        using (Document document = new Document(PageSize.A4)) //要寫PDF的文件，建構子沒填的話預設直式A4
                        {
                            PdfWriter pdfWriter = PdfWriter.GetInstance(document, output); //指定文件預設開檔時的縮放為100%
                            PdfDestination pdfDestination = new PdfDestination(PdfDestination.XYZ, 0, document.PageSize.Height, 1f); //開啟Document文件 

                            // 使用微軟正黑體
                            //BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\FONTS\\MSJHL.TTC,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            //Font font = new Font(baseFont);
                            // 3. 添加頁碼
                            CustomPageEventHandler pageEventHandler = new CustomPageEventHandler();
                            pageEventHandler.CompanyText = CompanyName;
                            pageEventHandler.CodeText = CodeText;
                            pdfWriter.PageEvent = pageEventHandler;

                            document.Open(); //開啟Document文件 
                            string 微軟正黑體 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "msjhl.ttc,0"); //微軟正黑體

                            BaseFont baseFont = BaseFont.CreateFont(微軟正黑體, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                            XMLWorkerHelper.GetInstance().ParseXHtml(pdfWriter, document, input, null, Encoding.UTF8, new UnicodeFontFactory1()); //使用XMLWorkerHelper把Html parse到PDF檔裡

                            PdfAction action = PdfAction.GotoLocalPage(1, pdfDestination, pdfWriter); //將pdfDest設定的資料寫到PDF檔

                            pdfWriter.SetOpenAction(action);

                        }
                    }

                    pdfFile = output.ToArray();
                }

                string fileGuid = Guid.NewGuid().ToString();
                Session[fileGuid] = pdfFile;


                #endregion

                #region //條碼圖片刪除
                DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/PdfTemplate/MES/CustomerComplaintCardImg"));
                FileInfo[] files1 = di.GetFiles();
                foreach (FileInfo file1 in files1)
                {
                    file1.Delete();
                }
                #endregion

                #endregion

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "OK",
                    fileGuid,
                    fileName = CompanyName + "客訴單據:" + CcNo,
                    fileExtension = ".pdf"
                });
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

        #region //DealWithContent 撈取單據階段資料用
        public string DealWithContent(string StageDataD2 = "", int Stage = -1,bool Appendix = false)
        {
            string retunStr = "";
            try
            {
                string Content = "";
                int row = 0;
                string[] splitStr1 = { "(^o^)" }; //自行設定切割字串
                string[] splitStr2 = { "(ToT)" }; //自行設定切割字串

                string[] result1 = StageDataD2.Split(splitStr1, StringSplitOptions.None);
                int countRow = -1;
                List<string> D7McListIndex = new List<string>();
                List<string> D7McListContent = new List<string>();
                foreach (var item in result1)
                {
                    if (item != "")
                    {
                        countRow++;
                        string McName = item.Split(splitStr2, StringSplitOptions.None)[0];
                        string content = item.Split(splitStr2, StringSplitOptions.None)[1];
                        string ConfirmUser = item.Split(splitStr2, StringSplitOptions.None)[3];
                        string ConfirmDate = item.Split(splitStr2, StringSplitOptions.None)[4];

                        row += content.Length / 28 + (content.Length % 28 != 0 ? 1 : 0);

                        int byteLength = Encoding.UTF8.GetByteCount(McName);
                        int fontSize = 14;
                        if (byteLength > 14) fontSize = 12;


                        Content += @"<tr>
                                            <td colspan=""4"" style=""width:16%;font-size:" + fontSize + @"px"">
                                                " + McName + @"
                                            </td>
                                            <td colspan=""15"" style=""width:60%;""><pre>" + content + @"</pre>
                                                
                                            </td>
                                            <td colspan=""3"" style=""width:12%;text-align:center;"">
                                                " + ConfirmUser + @"
                                            </td>
                                            <td colspan=""3"" style=""width:12%;font-size:12px;text-align:center;"">
                                                " + ConfirmDate + @"
                                            </td>
                                        </tr>";
                        if(Stage == 8 && Appendix == true)
                        {
                            Content = content;
                        }
                        //if (Stage != 7)
                        //{
                        //    row += content.Length / 28 + (content.Length % 28 != 0 ? 1 : 0);

                        //    int byteLength = Encoding.UTF8.GetByteCount(McName);
                        //    int fontSize = 14;
                        //    if (byteLength > 14) fontSize = 12;


                        //    Content += @"<tr>
                        //                    <td colspan=""4"" style=""width:16%;font-size:" + fontSize + @"px"">
                        //                        " + McName + @"
                        //                    </td>
                        //                    <td colspan=""15"" style=""width:60%;"">
                        //                        " + content + @"
                        //                    </td>
                        //                    <td colspan=""3"" style=""width:12%;text-align:center;"">
                        //                        " + ConfirmUser + @"
                        //                    </td>
                        //                    <td colspan=""3"" style=""width:12%;text-align:center;"">
                        //                        " + ConfirmDate + @"
                        //                    </td>
                        //                </tr>";
                        //}
                        //else if (Stage == 7)
                        //{
                        //    row += content.Length / 40 + (content.Length % 40 != 0 ? 1 : 0);

                        //    D7McListIndex.Add(McName);
                        //    D7McListContent.Add(content);
                        //}
                    }
                }

                //if (Stage == 7)
                //{
                //    #region //Request - 客訴單據卡D7資料
                //    JObject jsonResponseD7 = new JObject();
                //    string dataRequestD7 = customerComplaintDA.GetCCMenuCategory(-1, "D7", "", "", "", -1, -1);
                //    jsonResponseD7 = BaseHelper.DAResponse(dataRequestD7);
                //    dynamic[] dataD7 = JsonConvert.DeserializeObject<JObject[]>(JObject.Parse(dataRequestD7)["data"].ToString());
                //    string contentD7 = "";
                //    countRow = 0;
                //    row = row + 6 < 14 ? 14 : row;
                //    foreach (var item1 in dataD7)
                //    {
                //        string Color = "";

                //        string keyWord = item1.McName;
                //        int byteLength = Encoding.UTF8.GetByteCount(keyWord);
                //        int index = D7McListIndex.IndexOf(keyWord);
                //        int fontSize = 14;
                //        if (byteLength > 14) fontSize = 12;
                //        if (index != -1)
                //        {
                //            contentD7 = D7McListContent[index].ToString();
                //            Color = "black";
                //        }
                //        else
                //        {
                //            contentD7 = "";
                //        }

                //        Content += @"<tr>
                //                            <td colspan=""4"" style=""width:16%;font-size:"+ fontSize + @"px"">
                //                               <span style=""background-color:" + Color + @";"">口</span> " + keyWord + @"
                //                            </td>
                //                            <td colspan=""21"" style=""width:84%;"">
                //                                " + contentD7 + @"
                //                            </td>
                //                        </tr>";
                //    }
                    //#endregion
                //}

                retunStr = (row + countRow).ToString() + "♡" + Content;
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

            return retunStr;
        }
        #endregion

        #region //PDF輸出執行相關
        public class CustomPageEventHandler : PdfPageEventHelper
        {
            public string CompanyText { get; set; }
            public string CodeText { get; set; }
            public int Page { get; set; }
            public PdfTemplate tpl;
            public BaseFont bf;
            public override void OnOpenDocument(PdfWriter writer, Document document)
            {
                try
                {
                    tpl = writer.DirectContent.CreateTemplate(100, 100);
                    bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                }
                catch (Exception e)
                {

                    //throw new ExceptionConverter(e);
                }
            }
            public override void OnEndPage(PdfWriter writer, Document document)
            {
                base.OnEndPage(writer, document);

                // 添加頁碼、公司名稱和代碼
                PdfContentByte cb = writer.DirectContent;
                int pageN = writer.PageNumber;
                string pageText = "Page " + pageN;
                string companyText = CompanyText;
                string codeText = CodeText;
                Page = pageN;
                // 設置字體，使用微軟正黑體
                BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\msjh.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                float fontSize = 12;

                // 計算頁碼、公司名稱和代碼的位置
                float leftX = document.Left + 10; // 最左邊留10點的邊距
                float centerX = (document.Right + document.Left) / 2;
                float rightX = document.Right - 10; // 最右邊留10點的邊距
                float yPos = document.Bottom - 15; // 最下方留20點的邊距

                float pageX = leftX;
                float companyX = centerX - (baseFont.GetWidthPoint(companyText, fontSize) / 2);
                float codeX = rightX - baseFont.GetWidthPoint(codeText, fontSize);

                float len = bf.GetWidthPoint(pageText, 12);

                // 顯示頁碼、公司名稱和代碼
                cb.BeginText();
                cb.SetFontAndSize(baseFont, fontSize);
                cb.SetTextMatrix(pageX, yPos);
                cb.ShowText(pageText);
                cb.EndText();
                cb.AddTemplate(tpl, pageX + len, yPos);//定位“y页” 在具体的页面调试时候需要更改这xy的坐标

                cb.BeginText();
                cb.SetTextMatrix(companyX, yPos);
                cb.ShowText(CompanyText);
                cb.EndText();

                cb.BeginText();
                cb.SetTextMatrix(codeX, yPos);
                cb.ShowText(codeText);
                cb.EndText();
            }

            public override void OnCloseDocument(PdfWriter writer, Document document)
            {
                base.OnCloseDocument(writer, document);
                // 計算頁碼、公司名稱和代碼的位置

                float leftX = document.Left + 30; // 最左邊留10點的邊距
                float pageX = leftX;
                float centerX = (document.Right + document.Left) / 2;
                float rightX = document.Right - 10; // 最右邊留10點的邊距
                float yPos = document.Bottom - 15; // 最下方留20點的邊距

                int totalPageNumber = writer.PageNumber - 1; // 減去封面頁
                                                             // 在 PDF 文件的指定位置添加總頁數


                tpl.BeginText();
                tpl.SetFontAndSize(bf, 13);
                tpl.SetTextMatrix(0, 0);
                tpl.ShowText("/" + totalPageNumber);
                tpl.EndText();
            }
        }

        public class UnicodeFontFactory1 : FontFactoryImp
        {
            private static readonly string 微軟正黑體 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "msjhl.ttc,0"); //微軟正黑體

            public override Font GetFont(string fontname, string encoding, bool embedded, float size, int style, BaseColor color, bool cached)
            {
                //可用Arial或標楷體，自己選一個
                BaseFont baseFont = BaseFont.CreateFont(微軟正黑體, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                return new Font(baseFont, size, style, color);
            }
        }

        #endregion
        #endregion
    }
}