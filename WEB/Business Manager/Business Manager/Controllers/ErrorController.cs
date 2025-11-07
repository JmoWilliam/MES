using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ErrorController : WebController
    {
        #region //View
        public ActionResult NotFound()
        {
            this.Response.StatusCode = 404;

            return View();
        }

        public ActionResult InternalServerError()
        {
            this.Response.StatusCode = 500;

            return View();
        }
        #endregion
    }
}