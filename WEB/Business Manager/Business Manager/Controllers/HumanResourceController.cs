using Helpers;
using HRMDA;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class HumanResourceController : WebController
    {
        private HumanResourceDA humanResourceDA = new HumanResourceDA();

        #region //View
        #endregion

        #region //Get
        #endregion

        #region //Add
        #endregion

        #region //Update        
        #endregion

        #region //Delete
        #endregion

        #region //Api
        #region //UpdateDepartmentSynchronize -- 部門資料同步
        [HttpPost]
        [Route("api/HRM/DepartmentSynchronize")]
        public void UpdateDepartmentSynchronize(string Company, string SecretKey, string UpdateDate)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateDepartmentSynchronize");
                #endregion

                #region //Request
                dataRequest = humanResourceDA.UpdateDepartmentSynchronize(Company, UpdateDate);
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

        #region //UpdateUserSynchronize -- 使用者資料同步
        [HttpPost]
        [Route("api/HRM/UserSynchronize")]
        public void UpdateUserSynchronize(string Company, string SecretKey, string UpdateDate, string UserNo)
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "UpdateUserSynchronize");
                #endregion

                #region //Request
                dataRequest = humanResourceDA.UpdateUserSynchronize(Company, UpdateDate, UserNo);
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

        #region//GetDepartmentManagerNames --取得課、部、處主管姓名
        [HttpPost]
        [Route("api/HRM/DepartmentManagerNames")]
        public void GetDepartmentManagerNames(string Company="", string SecretKey = "", string UserNo = "", string DepNo = "")
        {
            try
            {
                #region //Api金鑰驗證
                ApiKeyVerify(Company, SecretKey, "GetDepartmentManagerNames");
                #endregion

                #region //Request
                dataRequest = humanResourceDA.GetDepartmentManagerNames(Company,UserNo, DepNo);
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