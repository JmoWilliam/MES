using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class UserController : WebController
    {
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

        public ActionResult MLogin()
        {
            return View();
        }
        #endregion

        #region //Get
        #region //GetLoginStatus 登入狀態資料
        [HttpPost]
        public void GetLoginStatus()
        {
            try
            {
                WebApiLoginCheck();

                var result = new
                {
                    linkType = BaseHelper.ClientLinkType(),
                    linkAddress = BaseHelper.ClientIP(),
                    linkName = BaseHelper.ClientComputer(),
                    test = Request.UserHostAddress
                };

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    result
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

        #region //GetCompanySwitchVerify 公司別切換驗證
        [HttpPost]
        public void GetCompanySwitchVerify()
        {
            try
            {
                WebApiLoginCheck("UserManagement", "company-switch");

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = ""
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

        #region //GetLoginByKey 取得登入者資訊(cookie)
        [HttpPost]
        public void GetLoginByKey(string UserNo = "", string KeyText = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                dataRequest = basicInformationDA.GetLoginByKey(UserNo, KeyText);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    if (JObject.Parse(dataRequest)["data"].ToString() == "[]")
                    {
                        dataRequest = basicInformationDA.DeleteUserLoginKey(Request.Cookies.Get("Login").Value, Request.Cookies.Get("LoginKey").Value);
                        Session.Clear();
                        Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                        Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
                    }
                    else
                    {
                        var returnData = new
                        {
                            UserId = Convert.ToInt32(result[0]["UserId"]),
                            UserNo = result[0]["UserNo"].ToString(),
                            UserName = result[0]["UserName"].ToString(),
                            CompanyName = result[0]["CompanyName"].ToString(),
                            CompanyNo = result[0]["CompanyNo"].ToString(),
                            CompanyId = Convert.ToInt32(result[0]["CompanyId"]),
                            DepartmentId = Convert.ToInt32(result[0]["DepartmentId"])
                        };

                        #region //判斷公司別是否手動
                        if (Session["CompanySwitch"] != null)
                        {
                            if (Session["CompanySwitch"].ToString() == "manual")
                            {
                                returnData = new
                                {
                                    UserId = Convert.ToInt32(result[0]["UserId"]),
                                    UserNo = result[0]["UserNo"].ToString(),
                                    UserName = result[0]["UserName"].ToString(),
                                    CompanyName = Session["CompanyName"].ToString(),
                                    CompanyNo = "",
                                    CompanyId = Convert.ToInt32(Session["CompanyId"]),
                                    DepartmentId = Convert.ToInt32(result[0]["DepartmentId"])
                                };
                            }
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            returnData
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

        #region//取得 Mamo 憑證
        [HttpPost]
        public void GetMamoLoginKey(string account = "", string password = "",string timestamp="")
        {
            try
            {
                string ip= BaseHelper.ClientIP();
                string MMQuery = "mm";
                string dataRequest = "";
                using (HttpClient client = new HttpClient()) {
                    string apiUrl = "http://192.168.20.46:2536/Mattermost/reset";
                    client.Timeout = TimeSpan.FromMinutes(60);
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl)) {
                        MultipartFormDataContent content = new MultipartFormDataContent
                        {
                            { new StringContent(account.ToString()), "account" },
                            { new StringContent(password.ToString()), "password" },
                            { new StringContent(timestamp.ToString()), "timestamp" },
                            { new StringContent(ip.ToString()), "ip" },
                            { new StringContent(MMQuery.ToString()), "MMQuery" }
                        };
                        httpRequestMessage.Content = content;
                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                dataRequest = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();                               
                                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                                {
                                    #region //Response
                                    jsonResponse = JObject.FromObject(new
                                    {
                                        status = "success",
                                        msg = ""
                                    });
                                    #endregion
                                }
                                else {
                                    #region //Response
                                    jsonResponse = JObject.FromObject(new
                                    {
                                        status = "error",
                                        msg = JObject.Parse(dataRequest)["message"].ToString()
                                    });
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("伺服器連線異常");
                            }
                        }
                    }
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

        #endregion

        #region //Add
        #endregion

        #region //Update
        #region //UpdatePassword 密碼修改
        [HttpPost]
        public void UpdatePassword(int UserId = -1, string NewPassword = "", string ConfirmPassword = "")
        {
            try
            {
                WebApiLoginCheck();

                #region //Request
                if (UserId <= 0) UserId = Convert.ToInt32(Session["UserId"]);

                dataRequest = basicInformationDA.UpdatePassword(UserId, NewPassword, ConfirmPassword);
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

        #region //UpdatePasswordReset 密碼重置
        [HttpPost]
        public void UpdatePasswordReset(int UserId = -1)
        {
            try
            {
                WebApiLoginCheck("UserManagement", "password-reset");

                #region //Request
                dataRequest = basicInformationDA.UpdatePasswordReset(UserId);
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

        #region //UpdateCompanySwitch 公司別切換
        [HttpPost]
        public void UpdateCompanySwitch(int CompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("UserManagement", "company-switch");

                #region //Request
                dataRequest = basicInformationDA.GetCompany(CompanyId, "", "", "A", "", -1, -1);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    Session["CompanyId"] = Convert.ToInt32(result[0]["CompanyId"]);
                    Session["CompanyName"] = result[0]["CompanyName"].ToString();
                    Session["CompanySwitch"] = "manual";
                }

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
        #endregion

        #region //Custom
        #region //登入
        [HttpPost]
        public void Login(string Account, string Password)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetLogin(Account, Password);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    string passwordStatus = result[0]["PasswordStatus"].ToString();

                    if (passwordStatus == "N")
                    {
                        if (Convert.ToDateTime(result[0]["PasswordExpire"]) > DateTime.Now)
                        {
                            LoginLog(Convert.ToInt32(result[0]["UserId"]), Account);

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
                            string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Account), 12, 32);

                            Session["NewLoginIV"] = iv;
                            Session["NewLogin"] = BaseHelper.AESEncrypt(Account, key, iv);
                            Session.Timeout = 300;

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "expirePassword",
                                msg = "密碼過期"
                            });
                            #endregion
                        }
                    }
                    else
                    {
                        string iv = BaseHelper.RandomCode(16);
                        string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Account), 12, 32);

                        Session["NewLoginIV"] = iv;
                        Session["NewLogin"] = BaseHelper.AESEncrypt(Account, key, iv);
                        Session.Timeout = 300;

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "newPassword",
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

        #region //新密碼登入
        [HttpPost]
        public void NewLogin(string Account, string NewPassword, string ConfirmPassword)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.UpdateNewLogin(Account, NewPassword, ConfirmPassword);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];

                    LoginLog(Convert.ToInt32(result[0]["UserId"]), Account);

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

        #region //登出
        public void Logout()
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.DeleteUserLoginKey(Request.Cookies.Get("Login").Value, Request.Cookies.Get("LoginKey").Value);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion

                Session.Clear();
                Response.Cookies["Login"].Expires = DateTime.Now.AddDays(-1);
                Response.Cookies["LoginKey"].Expires = DateTime.Now.AddDays(-1);
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

        #region //忘記密碼
        #region //MAMO推播密鑰網址
        [HttpPost]
        public void SendResetMail(string UserNo)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.UpdateSendResetMail(UserNo);
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

        #region //重製密碼登入
        [HttpPost]
        public void ResetPassword(string Token, string Password, string ConfirmPassword)
        {
            try
            {
                if (Password != ConfirmPassword) throw new SystemException("密碼與確認密碼不同!!");

                #region //Request
                dataRequest = basicInformationDA.UpdateResetPassword(Token, Password);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest);

                    #region //更改MAMO密碼
                    string timestamp = (DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() / 1000).ToString();
                    JObject mamoDataRequest = UpdateMamoPassword(result["UserNo"].ToString(), Password, timestamp);
                    if (mamoDataRequest["status"].ToString() != "success")
                    {
                        throw new SystemException(mamoDataRequest["msg"].ToString());
                    }
                    #endregion

                    LoginLog(Convert.ToInt32(result["UserId"]), result["UserNo"].ToString());

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

        #region //更改MAMO密碼
        public JObject UpdateMamoPassword(string account = "", string password = "", string timestamp = "")
        {
            try
            {
                string ip = BaseHelper.ClientIP();
                string MMQuery = "mm";
                string dataRequest = "";
                using (HttpClient client = new HttpClient())
                {
                    string apiUrl = "http://192.168.20.46:2536/Mattermost/reset";
                    client.Timeout = TimeSpan.FromMinutes(60);
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent content = new MultipartFormDataContent
                        {
                            { new StringContent(account.ToString()), "account" },
                            { new StringContent(password.ToString()), "password" },
                            { new StringContent(timestamp.ToString()), "timestamp" },
                            { new StringContent(ip.ToString()), "ip" },
                            { new StringContent(MMQuery.ToString()), "MMQuery" }
                        };
                        httpRequestMessage.Content = content;
                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                dataRequest = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                                if (JObject.Parse(dataRequest)["status"].ToString() == "success")
                                {
                                    #region //Response
                                    jsonResponse = JObject.FromObject(new
                                    {
                                        status = "success",
                                        msg = ""
                                    });
                                    #endregion
                                }
                                else
                                {
                                    #region //Response
                                    jsonResponse = JObject.FromObject(new
                                    {
                                        status = "error",
                                        msg = JObject.Parse(dataRequest)["message"].ToString()
                                    });
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("伺服器連線異常");
                            }
                        }
                    }
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

            return jsonResponse;
        }
        #endregion

        #endregion
        #endregion
    }
}