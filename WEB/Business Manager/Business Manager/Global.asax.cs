using Business_Manager.Controllers;
using Helpers;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace Business_Manager
{
    public class MvcApplication : HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            RouteConfig.RegisterRoutes(RouteTable.Routes);
        }

        protected void Application_Error()
        {
            var exception = Server.GetLastError();
            Response.Clear();
            Server.ClearError();
            var routeData = new RouteData();
            routeData.Values["controller"] = "Error";
            routeData.Values["action"] = "InternalServerError";
            routeData.Values["exception"] = exception;
            Response.StatusCode = 500;
            if (exception is HttpException httpException)
            {
                Response.StatusCode = httpException.GetHttpCode();
                switch (Response.StatusCode)
                {
                    case 404:
                        routeData.Values["action"] = "NotFound";
                        break;
                }
            }

            IController errorsController = new ErrorController();
            var rc = new RequestContext(new HttpContextWrapper(Context), routeData);
            errorsController.Execute(rc);
        }

        protected void Session_Start()
        {
            Session["CompanyId"] = -1;
            Session["CompanyName"] = "";
            Session["CompanySwitch"] = "auto";
        }
    }
}
