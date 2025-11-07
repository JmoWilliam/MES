using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BPMDA;

namespace Business_Manager.Controllers
{
    public class BPMController : WebController
    {
        private PLMBudgetDA pLMBudgetDA = new PLMBudgetDA();

        #region//專案預算BPM拋轉
        [HttpPost]
        [Route("api/BPM/apiPLMBudgetBPM")]
        public void ApiPLMBudgetBPM(string CompanyNo = "",string UserNo="", string PLMBudget = "", int ProjectType = -1)
        {
            try
            {
                #region //Request
                pLMBudgetDA = new PLMBudgetDA();
                dataRequest = pLMBudgetDA.UpdatePLMBudgetTransferBpm(CompanyNo, UserNo, PLMBudget, ProjectType);
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

    }
}