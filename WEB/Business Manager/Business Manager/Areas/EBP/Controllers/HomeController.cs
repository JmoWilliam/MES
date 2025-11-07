using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Areas.EBP.Controllers
{
    public class HomeController : Controller
    {
        #region //View
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Billboard()
        {
            return View();
        }
        public ActionResult Billboard2()
        {
            return View();
        }
        public ActionResult Activity()
        {
            return View();
        }
        public ActionResult Club()
        {
            return View();
        }
        #endregion
    }
}