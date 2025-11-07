using System.Web.Mvc;

namespace Business_Manager.Areas.IM
{
    public class IMAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "IM";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                name: "IM_default",
                url: "IM/{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}