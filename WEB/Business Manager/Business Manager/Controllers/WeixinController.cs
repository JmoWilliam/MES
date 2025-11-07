using Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class WeixinController : Controller
    {
        #region //微信測試
        [HttpGet]
        [Route("api/BAS/Weixin")]
        public object Weixin(string echostr, string signature, string timestamp, string nonce)
        {
            try
            {
                WeixinHelper weixinHelper = new WeixinHelper();

                if (weixinHelper.CheckSignature(signature, timestamp, nonce))
                {
                    return echostr;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
        #endregion

        [HttpPost]
        [Route("api/BAS/Weixin/ServerIP")]
        public object GetServerIP()
        {
            try
            {
                WeixinHelper weixinHelper = new WeixinHelper();

                return weixinHelper.GetServerIP();
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}