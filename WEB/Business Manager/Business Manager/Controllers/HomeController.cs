using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class HomeController : WebController
    {
        #region //View
        public ActionResult Index()
        {
            ViewLoginCheck();

            return RedirectToAction("Gate", "Home", null);
        }

        public ActionResult Gate()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult GateDetail()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion
    }
}