using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using EBPDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class MealController : WebController
    {
        private MealDA mealDA = new MealDA();

        #region //View
        public ActionResult MealManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetRestaurantMeal 取得餐點資料
        [HttpPost]
        public void GetRestaurantMeal(int MealId = -1, int RestaurantId = -1, string RestaurantName = "", string MealName = "", double MealPrice = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetRestaurantMeal(MealId, RestaurantId, RestaurantName, MealName, MealPrice, Status
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

        #region //GetRestaurantId 取得餐廳名稱(Cmb用)
        [HttpPost]
        public void GetRestaurantId(int RestaurantId = -1, string RestaurantNo = "", string RestaurantName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetRestaurantId(RestaurantId, RestaurantNo, RestaurantName
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

        #region //GetMealCategory 取得餐點類別資料
        [HttpPost]
        public void GetMealCategory(int MealCgId = -1, int MealId = -1, int RestaurantId = -1, string CategoryName = "", string CategoryType = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetMealCategory(MealCgId, MealId, RestaurantId, CategoryName, CategoryType
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

        #region //GetMcDetail 取得餐點選項(細節)資料
        [HttpPost]
        public void GetMcDetail(int McDetailId = -1, int MealCgId = -1, int MealId = -1, string CategoryName = "", string McDetailName = ""
            , double McDetailPrice = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetMcDetail(McDetailId, MealCgId, MealId, CategoryName, McDetailName
                    , McDetailPrice
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

        #region //GetMealCgId 取得類別名稱(Cmb用)
        [HttpPost]
        public void GetMealCgId(int MealId = -1, int MealCgId = -1, string CategoryName = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetMealCgId(MealId, MealCgId, CategoryName
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

        #region //GetDailyMeal 取得每日餐點資料
        [HttpPost]
        public void GetDailyMeal(int MealId = -1, string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetDailyMeal(MealId, StartDate, EndDate);
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

        #region //GetLotDailyMeal 取得每日餐點資料(批次)
        [HttpPost]
        public void GetLotDailyMeal(string StartDate = "", string EndDate = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetLotDailyMeal(StartDate, EndDate);
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

        #region //GetMealId 取得餐點資料(Cmb用)
        [HttpPost]
        public void GetMealId()
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read,constrained-data");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetMealId();
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

        #region //GetCalendarDate 取得訂餐開放時段資料
        [HttpPost]
        public void GetCalendarDate(string OrderType = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "read");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.GetCalendarDate(OrderType);
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
        #region //AddRestaurantMeal 餐點資料新增
        [HttpPost]
        public void AddRestaurantMeal(int RestaurantId = -1, string MealName = "", double MealPrice = -1, int MealImage = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "add");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.AddRestaurantMeal(RestaurantId, MealName, MealPrice, MealImage);
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

        #region //AddMealCategory 新增餐點類別資料
        [HttpPost]
        public void AddMealCategory(int MealId = -1, int RestaurantId = -1, string CategoryName = "", string CategoryType = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "add");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.AddMealCategory(MealId, RestaurantId, CategoryName, CategoryType);
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

        #region //AddMcDetail 新增餐點選項(細節)資料
        [HttpPost]
        public void AddMcDetail(int MealCgId = -1, string McDetailName = "", double McDetailPrice = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "add");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.AddMcDetail(MealCgId, McDetailName, McDetailPrice);
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

        #region //AddDailyMeal 新增每日餐點資料
        [HttpPost]
        public void AddDailyMeal(int CalendarId = -1, string Meals = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "add");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.AddDailyMeal(CalendarId, Meals);
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

        #region //AddLotDailyMeal 新增每日餐點資料(批次)
        [HttpPost]
        public void AddLotDailyMeal(string Calendars = "", string Meals = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "add");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.AddLotDailyMeal(Calendars, Meals);
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

        #region //AddCopyRestaurantMeal 複製餐點資料
        [HttpPost]
        public void AddCopyRestaurantMeal(int CompanyId = -1, int MealId = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "add");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.AddCopyRestaurantMeal(CompanyId, MealId);
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
        #region //UpdateRestaurantMeal 餐點資料更新
        [HttpPost]
        public void UpdateRestaurantMeal(int MealId = -1, int RestaurantId = -1, string MealName = "", double MealPrice = -1, int MealImage = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "update");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.UpdateRestaurantMeal(MealId, RestaurantId, MealName, MealPrice, MealImage);
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

        #region //UpdateMealStatus 餐點狀態更新
        [HttpPost]
        public void UpdateMealStatus(int MealId = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "status-switch");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.UpdateMealStatus(MealId);
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

        #region //UpdateMealCategory 更新餐點類別資料
        [HttpPost]
        public void UpdateMealCategory(int MealCgId = -1, int MealId = -1, int RestaurantId = -1, string CategoryName = "", string CategoryType = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "update");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.UpdateMealCategory(MealCgId, MealId, RestaurantId, CategoryName, CategoryType);
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

        #region //UpdateMcDetail 更新餐點選項(細節)資料
        [HttpPost]
        public void UpdateMcDetail(int McDetailId = -1, int MealCgId = -1, string McDetailName = "", double McDetailPrice = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "update");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.UpdateMcDetail(McDetailId, MealCgId, McDetailName, McDetailPrice);
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

        #region //UpdateCalendarDate 訂餐開放時段資料更新
        [HttpPost]
        public void UpdateCalendarDate(string StartTime = "", string EndTime = "", string OrderType = "")
        {
            try
            {
                WebApiLoginCheck("MealManagement", "calendarDate");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.UpdateCalendarDate(StartTime, EndTime, OrderType);
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
        #region //DeleteRestaurantMeal -- 刪除餐點資料
        [HttpPost]
        public void DeleteRestaurantMeal(int MealId = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "delete");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.DeleteRestaurantMeal(MealId);
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

        #region //DeleteMealCategory -- 刪除餐點類別資料
        [HttpPost]
        public void DeleteMealCategory(int MealCgId = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "delete");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.DeleteMealCategory(MealCgId);
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

        #region //DeleteMcDetail -- 刪除餐點選項(細節)資料
        [HttpPost]
        public void DeleteMcDetail(int McDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "delete");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.DeleteMcDetail(McDetailId);
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

        #region //DeleteDailyMeal 刪除每日餐點資料
        [HttpPost]
        public void DeleteDailyMeal(int CalendarId = -1, int MealId = -1)
        {
            try
            {
                WebApiLoginCheck("MealManagement", "delete");

                #region //Request
                mealDA = new MealDA();
                dataRequest = mealDA.DeleteDailyMeal(CalendarId, MealId);
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