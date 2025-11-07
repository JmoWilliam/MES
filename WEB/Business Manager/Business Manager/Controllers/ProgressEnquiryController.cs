using Helpers;
using Newtonsoft.Json.Linq;
using PDMDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;



namespace Business_Manager.Controllers
{
    public class ProgressEnquiryController : WebController
    {
        //private ProgressEnquiryDA progressEnquiryDA = new ProgressEnquiryDA();
        #region //View
        public ActionResult OrderProgress()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //View
        public ActionResult ProductionProgress()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion
    }
}