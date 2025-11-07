using System.Web.Mvc;

namespace Business_Manager.Areas.EBP
{
    public class EBPAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "EBP";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                name: "EBP_default",
                url: "EBP/{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}