using Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class MesPlatformController : WebController
    {
        #region //View
        public ActionResult ProductionHistory()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult ProductionHistoryForSupplier()
        {
            //ViewLoginCheck();

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
    }
}