using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class EbpPlatformController : WebController
    {
        #region //View
        public ActionResult FoodOrdering()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion
    }
}