using Helpers;
using Newtonsoft.Json.Linq;
using PDMDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class RdWorkManagementTestController : WebController
    {
        #region //View
        public ActionResult RdWorkBomTest()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SortTest()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion
    }
}