using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Areas.IM.Controllers
{
    public class HomeController : WebController
    {
        #region //View
        public ActionResult Index()
        {
            ViewIMLoginCheck();

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