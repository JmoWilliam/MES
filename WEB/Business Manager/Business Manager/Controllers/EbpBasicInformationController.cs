using EBPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class EbpBasicInformationController : WebController
    {
        private EbpBasicInformationDA ebpBasicInformationDA = new EbpBasicInformationDA();

        #region //View
        public ActionResult BoardType()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult Annual()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult ClubJob()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult RestaurantManagement()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult OrderMealSteward()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetBoardType 取得公告類別資料
        [HttpPost]
        public void GetBoardType(int BoardTypeId = -1, string TypeName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("BoardType", "read,constrained-data");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetBoardType(BoardTypeId, TypeName, Status
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

        #region //GetAnnual 取得年份資料
        [HttpPost]
        public void GetAnnual(int AnnualId = -1, int Annual = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("Annual", "read,constrained-data");

                #region //
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetAnnual(AnnualId, Annual, Status
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

        #region //GetClubJob 取得社團職位資料
        [HttpPost]
        public void GetClubJob(int ClubJobId = -1, string JobName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ClubJob", "read,constrained-data");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetClubJob(ClubJobId, JobName, Status
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

        #region //GetRestaurant 取得餐廳資料
        [HttpPost]
        public void GetRestaurant(int RestaurantId = -1, string RestaurantNo = "", string RestaurantName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "read");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetRestaurant(RestaurantId, RestaurantNo, RestaurantName, Status
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

        #region//GetRestaurantFile 取得餐廳上傳檔案資料
        [HttpPost]
        public void GetRestaurantFile(int RestaurantId = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "read");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetRestaurantFile(RestaurantId);
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

        #region //GetStewardInfo 取得事務人員資料
        [HttpPost]
        public void GetStewardInfo(int StewardId = -1, int UserId = -1, string UserName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OrderMealSteward", "read");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetStewardInfo(StewardId, UserId, UserName
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

        #region //GetStaffInfo 取得每日餐點資料
        [HttpPost]
        public void GetStaffInfo(int UserId = -1, int StaffUserId = -1, string UserName = "")
        {
            try
            {
                WebApiLoginCheck("OrderMealSteward", "read");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.GetStaffInfo(UserId, StaffUserId, UserName);
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
        #region //AddBoardType 公告類別資料新增
        [HttpPost]
        public void AddBoardType(string TypeName = "")
        {
            try
            {
                WebApiLoginCheck("BoardType", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddBoardType(TypeName);
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

        #region //AddAnnual 年份資料新增
        [HttpPost]
        public void AddAnnual(int Annual = -1)
        {
            try
            {
                WebApiLoginCheck("Annual", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddAnnual(Annual);
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

        #region //AddClubJob 社團職位資料新增
        [HttpPost]
        public void AddClubJob(string JobName = "")
        {
            try
            {
                WebApiLoginCheck("ClubJob", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddClubJob(JobName);
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

        #region //AddRestaurant 餐廳資料新增
        [HttpPost]
        public void AddRestaurant(string RestaurantNo = "", string RestaurantName = ""
            , string GuiNumber = "", string ResponsiblePerson = "", string Contact = ""
            , string TelNo = "", string FaxNo = "", string Email = "", string Address = ""
            , string AccountDay = "", string PaymentType = "", string InvoiceCount = ""
            , string RemitBank = "", string RemitAccount = "", string RestaurantRemark = ""
            , string StartTime = "", string EndTime = "")
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddRestaurant(RestaurantNo, RestaurantName
                    , GuiNumber, ResponsiblePerson, Contact
                    , TelNo, FaxNo, Email, Address
                    , AccountDay, PaymentType, InvoiceCount
                    , RemitBank, RemitAccount, RestaurantRemark
                    , StartTime, EndTime);
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

        #region //AddRestaurantFile 餐廳檔案新增
        [HttpPost]
        public void AddRestaurantFile(int RestaurantId = -1, string FileList = "", int FileDoc = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddRestaurantFile(RestaurantId, FileList, FileDoc);
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

        #region //AddRestaurantMeal 餐廳餐點新增
        [HttpPost]
        public void AddRestaurantMeal(int RestaurantId = -1, string MealList = "", int MealImage = -1, double MealPrice = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddRestaurantMeal(RestaurantId, MealList, MealImage, MealPrice);
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

        #region //AddOrderMealSteward 新增事務員相關資料
        [HttpPost]
        public void AddOrderMealSteward(int UserId = -1, string Staffs = "")
        {
            try
            {
                WebApiLoginCheck("OrderMealSteward", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddOrderMealSteward(UserId, Staffs);
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

        #region //AddCopyRestaurant 複製餐廳資料
        [HttpPost]
        public void AddCopyRestaurant(int CompanyId = -1, int RestaurantId = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "add");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.AddCopyRestaurant(CompanyId, RestaurantId);
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
        #region //UpdateBoardType 公告類別資料更新
        [HttpPost]
        public void UpdateBoardType(int BoardTypeId = -1, string TypeName = "")
        {
            try
            {
                WebApiLoginCheck("BoardType", "update");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateBoardType(BoardTypeId, TypeName);
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

        #region //UpdateBoardTypeStatus 公告類別狀態更新
        [HttpPost]
        public void UpdateBoardTypeStatus(int BoardTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("BoardType", "status-switch");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateBoardTypeStatus(BoardTypeId);
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

        #region //UpdateAnnual 公告類別資料更新
        [HttpPost]
        public void UpdateAnnual(int AnnualId = -1, int Annual = -1)
        {
            try
            {
                WebApiLoginCheck("Annual", "update");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateAnnual(AnnualId, Annual);
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

        #region //UpdateAnnualStatus 公告類別狀態更新
        [HttpPost]
        public void UpdateAnnualStatus(int AnnualId = -1)
        {
            try
            {
                WebApiLoginCheck("Annual", "status-switch");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateAnnualStatus(AnnualId);
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

        #region //UpdateClubJob 社團職位資料更新
        [HttpPost]
        public void UpdateClubJob(int ClubJobId = -1, string JobName = "")
        {
            try
            {
                WebApiLoginCheck("ClubJob", "update");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateClubJob(ClubJobId, JobName);
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

        #region //UpdateClubJobStatus 社團職位狀態更新
        [HttpPost]
        public void UpdateClubJobStatus(int ClubJobId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubJob", "status-switch");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateClubJobStatus(ClubJobId);
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

        #region //UpdateRestaurant 餐廳資料更新
        [HttpPost]
        public void UpdateRestaurant(int RestaurantId = -1, string RestaurantNo = "", string RestaurantName = ""
            , string GuiNumber = "", string ResponsiblePerson = "", string Contact = ""
            , string TelNo = "", string FaxNo = "", string Email = "", string Address = ""
            , string AccountDay = "", string PaymentType = "", string InvoiceCount = ""
            , string RemitBank = "", string RemitAccount = "", string RestaurantRemark = ""
            , string StartTime = "", string EndTime = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "update");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateRestaurant(RestaurantId, RestaurantNo, RestaurantName
                    , GuiNumber, ResponsiblePerson, Contact
                    , TelNo, FaxNo, Email, Address
                    , AccountDay, PaymentType, InvoiceCount
                    , RemitBank, RemitAccount, RestaurantRemark
                    , StartTime, EndTime, Status);
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

        #region //UpdateRestaurantFile 餐廳檔案更新
        [HttpPost]
        public void UpdateRestaurantFile(int RestaurantId = -1, string FileList = "", int FileDoc = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "update");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateRestaurantFile(RestaurantId, FileList, FileDoc);
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

        #region //UpdateRestaurantMeal 餐廳餐點更新
        [HttpPost]
        public void UpdateRestaurantMeal(int RestaurantId = -1, string MealList = "", int MealImage = -1, double MealPrice = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "update");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateRestaurantMeal(RestaurantId, MealList, MealImage, MealPrice);
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

        #region //UpdateRestaurantStatus 餐廳狀態更新
        [HttpPost]
        public void UpdateRestaurantStatus(int RestaurantId = -1)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "status-switch");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.UpdateRestaurantStatus(RestaurantId);
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
        #region //DeleteBoardType 公告類別資料刪除
        [HttpPost]
        public void DeleteBoardType(int BoardTypeId = -1)
        {
            try
            {
                WebApiLoginCheck("BoardType", "delete");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.DeleteBoardType(BoardTypeId);
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

        #region //DeleteAnnual 年份資料刪除
        [HttpPost]
        public void DeleteAnnual(int AnnualId = -1)
        {
            try
            {
                WebApiLoginCheck("Annual", "delete");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.DeleteAnnual(AnnualId);
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

        #region //DeleteClubJob 社團職位資料刪除
        [HttpPost]
        public void DeleteClubJob(int ClubJobId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubJob", "delete");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.DeleteClubJob(ClubJobId);
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

        #region//DeleteRestaurantFile --餐廳刪除上傳檔案資料
        [HttpPost]
        public void DeleteRestaurantFile(int FileDoc)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "delete");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.DeleteRestaurantFile(FileDoc);
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

        #region//DeleteRestaurantMeal --餐廳刪除上傳菜單資料
        [HttpPost]
        public void DeleteRestaurantMeal(int MealId)
        {
            try
            {
                WebApiLoginCheck("RestaurantManagement", "delete");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.DeleteRestaurantMeal(MealId);
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

        #region //DeleteStaffUser 刪除該事務下所屬同仁資料
        [HttpPost]
        public void DeleteStaffUser(int UserId = -1, int StaffUserId = -1)
        {
            try
            {
                WebApiLoginCheck("OrderMealSteward", "delete");

                #region //Request
                ebpBasicInformationDA = new EbpBasicInformationDA();
                dataRequest = ebpBasicInformationDA.DeleteStaffUser(UserId, StaffUserId);
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