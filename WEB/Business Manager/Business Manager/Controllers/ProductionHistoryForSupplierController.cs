using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ProductionHistoryForSupplierController : WebController
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult MesScanningInterfaceSupplier()
        {
            return View();
        }
    }
}