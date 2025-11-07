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
    public class ActivityController : WebController
    {
        private ActivityDA activityDA = new ActivityDA();

        #region //View
        public ActionResult AwardsKanban()
        {
            return View();
        }
        #endregion

        #region //Get
        #region //GetAwardsKanban 取得獎項看板資料
        [HttpPost]
        public void GetAwardsKanban(int ActivityId = -1, int AwardsId = -1)
        {
            try
            {
                #region //Request
                activityDA = new ActivityDA();
                dataRequest = activityDA.GetAwardsKanban(ActivityId, AwardsId);
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