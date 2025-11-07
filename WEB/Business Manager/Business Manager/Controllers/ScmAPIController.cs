using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Web.Mvc;
using System.Web.Hosting;
using SCMDA;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Reflection;
using Dapper;

namespace Business_Manager.Controllers
{
    public class ScmAPIController : WebController
    {
        private ScmApiDA scmApiDA = new ScmApiDA();

        #region //API
        #region//ConfirmQuotation 核准核價單(非正式核單邏輯) -- Luca 2024-04-03
        [HttpPost]
        [Route("api/SRM/ConfirmQuotation")]
        public void ConfirmQuotation(string QoErpPrefix = "", string QoErpNo = "", string CompanyNo = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "ConfirmQuotation");
                #endregion

                #region //Request
                scmApiDA = new ScmApiDA();
                dataRequest = scmApiDA.ConfirmQuotation(QoErpPrefix, QoErpNo, CompanyNo);
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

        #region //ConfirmPurchaseOrder 核准採購單
        [HttpPost]
        [Route("api/SRM/ConfirmPurchaseOrder")]
        public void ConfirmPurchaseOrder(string PoErpPrefix = "", string PoErpNo = "", string CompanyNo = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "ConfirmPurchaseOrder");
                #endregion

                #region //Request
                scmApiDA = new ScmApiDA();
                dataRequest = scmApiDA.ConfirmPurchaseOrder(PoErpPrefix, PoErpNo, CompanyNo);
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

        #region //ConfirmPurchaseOrderChange 核准採購變更單
        [HttpPost]
        [Route("api/SRM/ConfirmPurchaseOrderChange")]
        public void ConfirmPurchaseOrderChange(string PoErpPrefix = "", string PoErpNo = "", string Edition = "", string CompanyNo = "", string SecretKey = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(CompanyNo, SecretKey, "ConfirmPurchaseOrderChange");
                #endregion

                #region //Request
                scmApiDA = new ScmApiDA();
                dataRequest = scmApiDA.ConfirmPurchaseOrderChange(PoErpPrefix, PoErpNo, Edition, CompanyNo);
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