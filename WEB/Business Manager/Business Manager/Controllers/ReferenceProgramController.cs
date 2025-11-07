using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ReferenceProgramController : WebController
    {
        #region //View
        public ActionResult PopView()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion
    }
}