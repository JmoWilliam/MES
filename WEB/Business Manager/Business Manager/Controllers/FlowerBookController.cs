using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class FlowerBookController : WebController
    {
        // GET: FlowerBook
        public ActionResult RDCycleProcess()
        {
            //中揚光電研發循環流程圖
            return View();
        }

        public ActionResult SOCycleProcess()
        {
            //中揚光電銷售及收款循環作業流程圖
            return View();
        }

        public ActionResult PLMCycleProcess()
        {
            //中揚光電 - 整合產品開發流程圖
            return View();
        }

        public ActionResult MESCycleProcess()
        {
            //中揚光電 - 生產循環完整流程圖 (文件編號: A-B3CO)
            return View();
        }

        public ActionResult PouCycleProcess()
        {
            //中揚光電 - 採購及付款循環-01
            return View();
        }
        public ActionResult RFItoolCycleProcess()
        {
            //中揚光電 - RFI文件程序軟體分析流程圖<
            return View();
        }
        public ActionResult FourBigCycleProcess()
        {
            //中揚光電 - 四大循環
            return View();
        }
        public ActionResult ManufacturingUnitExecutionFlowchart()
        {
            //中揚光電 - 製造單位生產執行流程圖
            return View();
        }
        public ActionResult TenDayQualityInsightt()
        {
            //中揚光電 - 10天品質管理資料收集專案
            return View();
        }
        public ActionResult ProjectBlueprintFutureReadyManufacturing()
        {
            //中揚光電 - 【製造&生管部門啟動會議議程】需求收集重點 (藍圖專案：迎向未來的製造)
            return View();
        }
        public ActionResult JMOsupplyHub()
        {
            //中揚光電 - 中揚光電(JMO)供應商平台介紹
            return View();
        }
    }
}