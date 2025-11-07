using EBPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Areas.EBP.Controllers
{
    public class FoodOrderingController : WebController
    {
        private FoodOrderingDA foodOrderingDA = new FoodOrderingDA();

        #region //View
        public ActionResult Login()
        {
            return View();
        }
        public ActionResult Index()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult McDetail()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult DiningCar()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult Account()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult Account2()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult LotAccount()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult CurrentHistory()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult NumerousMenu()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult NumerousOrder()
        {
            FoodOrderingLoginCheck();

            return View();
        }
        public ActionResult OrderHistory()
        {
            FoodOrderingLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetMealInfo 取得餐點資訊
        public void GetMealInfo(string MealDate = "", string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetMealInfo(MealDate, OrderBy, PageIndex, PageSize);
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

        #region //GetMcDetail 取得客製化項目資訊
        public void GetMcDetail(int MealId = -1, string MealDate = "")
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetMcDetail(MealId, MealDate);
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

        #region //GetStaffUserId 取得事務維護人員名稱(Cmb用)
        [HttpPost]
        public void GetStaffUserId()
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetStaffUserId();
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

        #region //GetOrderHistory 取得點餐紀錄資訊
        [HttpPost]
        public void GetOrderHistory(int UserId = -1, string UmoDate = "")
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetOrderHistory(UserId, UmoDate);
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

        #region //GetMealDate 取得每日餐點資料
        [HttpPost]
        public void GetMealDate()
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetMealDate();
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

        #region //GetUserOrderHistory 取得個人餐點資料
        [HttpPost]
        public void GetUserOrderHistory(string UserNo = "", string UmoDate = "")
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetUserOrderHistory(UserNo, UmoDate);
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

        #region //CheckWorkday 取得工作日資訊
        [HttpPost]
        public void CheckWorkday()
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.CheckWorkday();
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

        #region //GetOrderTime 取得個人餐點資料
        [HttpPost]
        public void GetOrderTime(string OrderType = "")
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.GetOrderTime(OrderType);
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
        #region //AddMealCart 購物車資料新增
        [HttpPost]
        public void AddMealCart(string MealCart = "")
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.AddMealCart(MealCart);
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

        #region //AddLotCart 批量購物車資料新增
        [HttpPost]
        public void AddLotCart(string LotCart = "")
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.AddLotCart(LotCart);
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
        #endregion

        #region //Delete
        #region //DeleteOrderHistory 單人點餐紀錄刪除
        [HttpPost]
        public void DeleteOrderHistory(int UmoId = -1)
        {
            try
            {
                FoodOrderingApiLoginCheck();

                #region //Request
                foodOrderingDA = new FoodOrderingDA();
                dataRequest = foodOrderingDA.DeleteOrderHistory(UmoId);
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

        #region //Custom
        #region //GetLogin 使用者登入
        [HttpPost]
        public void GetLogin(string UserNo, string Password)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetLogin(UserNo, Password);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    string passwordStatus = result[0]["PasswordStatus"].ToString();

                    if (passwordStatus == "N")
                    {
                        if (Convert.ToDateTime(result[0]["PasswordExpire"]) > DateTime.Now)
                        {
                            LoginLog(Convert.ToInt32(result[0]["UserId"]), UserNo, "IM");

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "登入成功!"
                            });
                            #endregion
                        }
                        else
                        {
                            string iv = BaseHelper.RandomCode(16);
                            string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(UserNo), 12, 32);

                            Session["NewLoginIV"] = iv;
                            Session["NewLogin"] = BaseHelper.AESEncrypt(UserNo, key, iv);
                            Session.Timeout = 300;

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "expirePassword",
                                userId = Convert.ToInt32(result[0]["UserId"]),
                                msg = "密碼過期"
                            });
                            #endregion
                        }
                    }
                    else
                    {
                        string iv = BaseHelper.RandomCode(16);
                        string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(UserNo), 12, 32);

                        Session["NewLoginIV"] = iv;
                        Session["NewLogin"] = BaseHelper.AESEncrypt(UserNo, key, iv);
                        Session.Timeout = 300;

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "newPassword",
                            userId = Convert.ToInt32(result[0]["UserId"]),
                            msg = "首次登入"
                        });
                        #endregion
                    }
                }
                else
                {
                    #region //Response
                    jsonResponse = BaseHelper.DAResponse(dataRequest);
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

        #region //View檢查登入
        [NonAction]
        public void FoodOrderingLoginCheck()
        {
            bool verify = LoginVerify("FoodOrdering");

            if (!verify)
            {
                string redirectUrl = "http://" + Request.Url.Authority + Request.RawUrl,
                    redirectLoginUrl = Url.Action("Login", "FoodOrdering", new { Area = "EBP" });

                var uri = new Uri(redirectUrl);
                var newQueryString = HttpUtility.ParseQueryString(uri.Query);

                string pagePathWithoutQueryString = uri.GetLeftPart(UriPartial.Path).Replace("http://", "").Replace(Request.Url.Authority, "");
                redirectUrl = newQueryString.Count > 0 ? String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString) : pagePathWithoutQueryString;
                if (redirectUrl != "/") redirectLoginUrl += "?preUrl=" + Uri.EscapeDataString(redirectUrl);

                HttpContext.Response.Redirect(redirectLoginUrl);
            }
        }
        #endregion

        #region //Api檢查登入
        [NonAction]
        public void FoodOrderingApiLoginCheck()
        {
            bool verify = LoginVerify("FoodOrdering");

            if (!verify)
            {
                throw new SystemException("系統已登出，請重新登入");
            }
        }
        #endregion
        #endregion
    }
}