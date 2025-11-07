using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ChatController : WebController
    {
        #region //View
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult FriendList()
        {
            return View();
        }

        public ActionResult ChatList()
        {
            return View();
        }
        #endregion

        #region //Api
        [HttpPost]
        [Route("api/Chat/Login")]
        public void GetUserLogin(string username, string password)
        {
            try
            {
                #region //Request
                dataRequest = basicInformationDA.GetLogin(username, password);
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
    }
}