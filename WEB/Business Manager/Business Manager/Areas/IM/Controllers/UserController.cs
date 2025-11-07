using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Areas.IM.Controllers
{
    public class UserController : WebController
    {
        #region //View
        public ActionResult Login()
        {
            return View();
        }
        #endregion

        #region //Get
        #region //GetLogin 使用者登入
        [HttpPost]
        public void GetLogin(string Username, string Password)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetLogin(Username, Password);
                #endregion

                if (JObject.Parse(dataRequest)["status"].ToString() != "errorForDA")
                {
                    var result = JObject.Parse(dataRequest)["data"];
                    string passwordStatus = result[0]["PasswordStatus"].ToString();

                    if (passwordStatus == "N")
                    {
                        if (Convert.ToDateTime(result[0]["PasswordExpire"]) > DateTime.Now)
                        {
                            LoginLog(Convert.ToInt32(result[0]["UserId"]), Username, "IM");

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
                            string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Username), 12, 32);

                            Session["NewLoginIV"] = iv;
                            Session["NewLogin"] = BaseHelper.AESEncrypt(Username, key, iv);
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
                        string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Username), 12, 32);

                        Session["NewLoginIV"] = iv;
                        Session["NewLogin"] = BaseHelper.AESEncrypt(Username, key, iv);
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
        #endregion
    }
}