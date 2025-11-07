using ClosedXML.Excel;
using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using MESDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SCMDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ScmReportController : WebController
    {
        private ScmReportDA scmReportDA = new ScmReportDA();

        #region //View
        public ActionResult RpqReport()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetRpqReport 取得RPQ電商詢價流程
        [HttpPost]
        public void GetRpqReport(int RfqId = -1, int RfqDetailId = -1, string RfqNo = "", int MemberId = -1, string MemberName = "", string AssemblyName = "", int ProductUseId = -1
            , int RfqProTypeId = -1, string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RpqReport", "read");

                #region //Request
                scmReportDA = new ScmReportDA();
                dataRequest = scmReportDA.GetRpqReport(RfqId, RfqDetailId, RfqNo, MemberId, MemberName, AssemblyName, ProductUseId
                , RfqProTypeId,  Status
                , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetSaleOrderProgress 取得訂單相關資料(手機版)
        [HttpPost]
        public void GetSaleOrderProgress(string SoErpPrefix = "", string SoFullNo = "", string SearchKey = "", string Customer = "", string Salesmen = ""
            , string ConfirmStatus = "", string ClosureStatus = "", string StartDate = "", string EndDate = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OrderProgress", "read");

                #region //Request
                scmReportDA = new ScmReportDA();
                dataRequest = scmReportDA.GetSaleOrderProgress(SoErpPrefix, SoFullNo, SearchKey, Customer, Salesmen
                , ConfirmStatus, ClosureStatus, StartDate, EndDate
                , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #region //GetProductSoDetail 取得訂單相關資料(手機版)
        [HttpPost]
        public void GetProductSoDetail(string SoFullNo = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("OrderProgress", "read");

                #region //Request
                scmReportDA = new ScmReportDA();
                dataRequest = scmReportDA.GetProductSoDetail(SoFullNo
                , OrderBy, PageIndex, PageSize);
                #endregion

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }
        #endregion

        #endregion
    }
}