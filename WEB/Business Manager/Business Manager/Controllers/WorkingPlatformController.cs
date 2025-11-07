using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class WorkingPlatformController : WebController
    {
        #region //View
        public ActionResult RandDWorkingPlatform()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult SalesWorkingPlatform()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult PCWorkingPlatform()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #endregion

        #region //Add
        #endregion

        #region //Update
        #endregion

        #region //Delete
        #endregion

        #region //Api
        #endregion
    }
}