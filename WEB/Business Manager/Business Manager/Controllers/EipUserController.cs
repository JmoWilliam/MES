using System;
using System.Web;
using System.Linq;
using System.Web.Mvc;
using System.Collections.Generic;
using Helpers;
using Newtonsoft.Json.Linq;
using BASDA;
using EIPDA;

namespace Business_Manager.Controllers
{
    public class EipUserController : WebController
    {
        private EipBasicInformationDA eipBasicInformationDA = new EipBasicInformationDA();
        private BasicInformationDA basicInformationDA = new BasicInformationDA();

        #region //View
        public ActionResult Login()
        {
            return View();
        }

        public ActionResult EditPassword()
        {
            ViewLoginCheck();

            return RedirectToAction("Gate", "Home", null);
        }

        #endregion

        #region //Custom
        #region //Login -- 登入 -- Chia Yuan -- 2023.7.7
        [HttpPost]
        [Route("api/EIP/Login")]
        public void Login(string Account, string Password)
        {
            try
            {
                HttpCookie MemberType = System.Web.HttpContext.Current.Request.Cookies.Get("MemberType");
                int memberType = -1;
                if (Account.IndexOf("@") > 0)
                {
                    #region //Request
                    dataRequest = eipBasicInformationDA.GetLogin(Account, Password);
                    #endregion

                    memberType = 1;
                }
                else
                {
                    #region //Request
                    dataRequest = basicInformationDA.GetLogin(Account, Password);
                    #endregion

                    memberType = 2;
                }

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    MemberType = new HttpCookie("MemberType")
                    {
                        Value = memberType.ToString(),
                        Expires = DateTime.Now.AddDays(1)
                    };
                    System.Web.HttpContext.Current.Response.Cookies.Add(MemberType);

                    var result = JObject.Parse(dataRequest)["data"];
                    string passwordStatus = result[0]["PasswordStatus"].ToString();

                    if (passwordStatus == "N")
                    {
                        //if (Convert.ToDateTime(result[0]["PasswordExpire"]) > DateTime.Now)
                        if (true)
                        {
                            JToken data = result.FirstOrDefault();
                            int MemberId = Convert.ToInt32(data["UserId"]);
                            string MemberName = data["UserName"].ToString();
                            string Address = memberType == 1? data["Address"].ToString() : string.Empty;
                            MemberType = HttpContext.Request.Cookies.Get("MemberType");

                            if (MemberType.Value == "1")
                            {
                                dataRequest = eipBasicInformationDA.UpdateUserLoginKey(MemberId, Account);
                            }
                            if (MemberType.Value == "2")
                            {
                                dataRequest = basicInformationDA.UpdateUserLoginKey(MemberId, Account);
                            }

                            HttpCookie Login = HttpContext.Request.Cookies.Get("Login");
                            HttpCookie LoginKey = HttpContext.Request.Cookies.Get("LoginKey");

                            //Session.Add("UserId", memberId);
                            //Session.Add("UserName", memberName);
                            //Session.Add("UserNo", Login.Value);
                            //Session.Add("LoginKey", LoginKey.Value);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                Login,
                                LoginKey,
                                MemberType,
                                MemberName,
                                Address,
                                Expires = Login.Expires.ToString("yyyy-MM-dd HH:mm:ss"),
                                status = "success"
                            });
                            #endregion
                        }
                        else
                        {
                            //密碼過期處理函式
                        }
                    }
                    else
                    {
                        //首次登入處理函式
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

        #region //Logout -- 登出 -- Chia Yuan -- 2023.7.7
        [HttpPost]
        [Route("api/EIP/Logout")]
        public void Logout(int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                if (MemberType == 1)
                {
                    #region //Request
                    dataRequest = eipBasicInformationDA.DeleteUserLoginKeyFromEIP(Account, KeyText);
                    #endregion
                }
                if (MemberType == 2)
                {
                    #region //Request
                    dataRequest = eipBasicInformationDA.DeleteUserLoginKeyFromBM(Account, KeyText);
                    #endregion
                }

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                ClearLoginStatus();
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

        #region //LoginVerifyFromEIP -- 登入狀態驗證 -- Chia Yuan -- 2023.7.7
        [NonAction]
        protected (bool verify, int UserId, string UserName, string UserNo) LoginVerifyFromEIP(string Account = "", string KeyText = "")
        {
            bool verify = false;
            int UserId = -1;
            string UserName = "", UserNo = "";

            if (!(string.IsNullOrWhiteSpace(Account) && string.IsNullOrWhiteSpace(KeyText)))
            {
                #region //檢查使用者登入IP
                #region //Request
                dataRequest = eipBasicInformationDA.GetLoginIPCheck(KeyText);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    if (result.Count() == 1)
                    {
                        for (int i = 0; i < result.Count(); i++)
                        {
                            verify = BaseHelper.ClientIP() == result[i]["LoginIP"].ToString();
                        }
                    }
                }
                #endregion

                #region //取得登入者資訊
                if (verify)
                {
                    #region //Request
                    dataRequest = eipBasicInformationDA.GetLoginByKey(Account, KeyText);
                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        if (result.Count() == 1)
                        {
                            JToken data = result.FirstOrDefault();
                            UserId = Convert.ToInt32(data["MemberId"]);
                            UserName = data["MemberName"].ToString();
                            UserNo = data["MemberEmail"].ToString();
                            //Session.Add("UserLock", true);
                            //Session.Add("UserId", Convert.ToInt32(data["MemberId"]));
                            //Session.Add("UserName", data["MemberName"].ToString());
                            //Session.Add("UserNo", data["MemberEmail"].ToString());
                            //Session.Add("LoginKey", data["KeyText"].ToString());
                        }
                        else
                            ClearLoginStatus();
                    }
                }
                else
                    ClearLoginStatus();
                #endregion
            }

            return (verify, UserId, UserName, UserNo);
        }
        #endregion

        #region //LoginVerifyFromBM -- 登入狀態驗證 -- Chia Yuan -- 2023.08-01
        [NonAction]
        protected (bool verify, int UserId, string UserName, string UserNo) LoginVerifyFromBM(string Account = "", string KeyText = "")
        {
            bool verify = false;
            int UserId = -1;
            string UserName = "", UserNo = "";

            if (!(string.IsNullOrWhiteSpace(Account) && string.IsNullOrWhiteSpace(KeyText)))
            {
                #region //檢查使用者登入IP
                #region //Request
                dataRequest = basicInformationDA.GetLoginIPCheck(Account);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    if (result.Count() == 1)
                    {
                        for (int i = 0; i < result.Count(); i++)
                        {
                            verify = BaseHelper.ClientIP() == result[i]["LoginIP"].ToString();
                        }
                    }
                }
                #endregion

                #region //取得登入者資訊
                if (verify)
                {
                    #region //Request
                    dataRequest = basicInformationDA.GetLoginByKey(Account, KeyText);
                    #endregion

                    if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                    {
                        var result = JObject.Parse(dataRequest)["data"];
                        if (result.Count() == 1)
                        {
                            JToken data = result.FirstOrDefault();
                            UserId = Convert.ToInt32(data["UserId"]);
                            UserName = data["UserName"].ToString();
                            UserNo = data["UserNo"].ToString();
                        }
                        else
                            ClearLoginStatus();
                    }
                }
                else
                    ClearLoginStatus();
                #endregion
            }

            return (verify, UserId, UserName, UserNo);
        }
        #endregion

        #region //ClearLoginStatus -- 清除登入狀態 -- Chia Yuan -- 2023.7.25
        /// <summary>
        /// 清除登入狀態
        /// </summary>
        protected void ClearLoginStatus()
        {
            Session.Clear();
            Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
            Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
            Response.Cookies["MemberType"].Expires = DateTime.Now.AddDays(-1);
        }
        #endregion

        #region  //ForgotPassword -- 寄送更換密碼通知信 -- Chia Yuan -- 2023.07.27
        [HttpPost]
        [Route("api/EIP/ForgotPassword")]
        public void ForgotPassword(string Account = "")
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.ForgotPassword(Account);
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

        #region //Get

        #region //GetLoginByKey -- 取得登入者資訊(前端傳入 cookie) -- Chia Yuan 2023.7.7
        [HttpPost]
        [Route("api/EIP/GetLoginByKey")]
        public void GetLoginByKey(int MemberType = -1, string Account = "", string KeyText = "")
        {
            try
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    data = new { }
                });

                if (MemberType == 1)
                {
                    (bool verify, int UserId, string UserName, string UserNo) = LoginVerifyFromEIP(Account, KeyText);
                    if (verify)
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = new { UserId, UserName }
                        });
                        #endregion
                    }
                    else
                        ClearLoginStatus();
                }
                else if (MemberType == 2)
                {
                    (bool verify, int UserId, string UserName, string UserNo) = LoginVerifyFromBM(Account, KeyText);
                    if (verify)
                    {
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = new { UserId, UserName }
                        });
                        #endregion
                    }
                    else
                        ClearLoginStatus();
                }
                else
                    ClearLoginStatus();
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
        #region //RegisterUser -- 使用者資料新增 -- Chia Yuan  -- 2023.7.7
        [HttpPost]
        [Route("api/EIP/RegisterUser")]
        public void RegisterUser(string Account = "", string Password = "", string MemberName = "", string OrgShortName = "", string ContactName = "", string ContactPhone = "")
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.RegisterUser(Account, Password, MemberName, OrgShortName, ContactName, ContactPhone);
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
        #region //UpdateMember -- 使用者資料更新 -- Chia Yuan  -- 2023.07.28
        [HttpPost]
        [Route("api/EIP/UpdateMember")]
        public void UpdateMember(string MemberName = "", string OrgShortName = "", string Address = "", string ContactName = "", string ContactPhone = ""
            , string Account = "", string KeyText = "")
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateMember(MemberName, OrgShortName, Address, ContactName, ContactPhone
                    , Account, KeyText);
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

        #region  //UpdatePasswordReset -- 重設密碼 -- Chia Yuan -- 2023.07.27
        [HttpPost]
        [Route("api/EIP/UpdatePasswordReset")]
        public void UpdatePasswordReset(string Password = "", string NewPassword = "", string KeyText = "")
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdatePasswordReset(Password, NewPassword, KeyText);
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

        #region  //UpdateMemberAccount -- 密碼更新 -- Chia Yuan -- 2023.07.27
        [HttpPost]
        [Route("api/EIP/UpdateMemberAccount")]
        public void UpdateMemberAccount(string Password = "", string NewPassword = "", string ConfirmPassword = ""
            , string Account = "", string KeyText = "")
        {
            try
            {
                #region //Request
                eipBasicInformationDA = new EipBasicInformationDA();
                dataRequest = eipBasicInformationDA.UpdateMemberAccount(Password, NewPassword, ConfirmPassword
                    , Account, KeyText);
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