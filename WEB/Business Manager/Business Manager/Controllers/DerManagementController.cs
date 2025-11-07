using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using FINDA;
using Helpers;
using Newtonsoft.Json.Linq;

namespace Business_Manager.Controllers
{
    public class DerManagementController : WebController
    {
        private DerManagementDA derManagementDA = new DerManagementDA();

        #region //View
        public ActionResult DerManagement()
        {
            ViewLoginCheck();
            return View();
        }
        #endregion

        #region //Get
        #region //GetDepartmentRate 取得部門費用率資料
        [HttpPost]
        public void GetDepartmentRate(int DepartmentRateId = -1, int DepartmentId = -1, int UserId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("DerManagement", "read");

                #region //Request
                derManagementDA = new DerManagementDA();
                dataRequest = derManagementDA.GetDepartmentRate(DepartmentRateId, DepartmentId, UserId
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

        #region //Add
        #region //AddDepartmentRate 新增部門費用率資料
        [HttpPost]
        public void AddDepartmentRate(int DepartmentId = -1, double ResourceRate = -1, double OverheadRate = -1)
        {
            try
            {
                WebApiLoginCheck("DerManagement", "add");

                #region //Request
                derManagementDA = new DerManagementDA();
                dataRequest = derManagementDA.AddDepartmentRate(DepartmentId, ResourceRate, OverheadRate);
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

        #region //Update
        #region //UpdateDepartmentRate 更新部門費用率資料
        [HttpPost]
        public void UpdateDepartmentRate(int DepartmentRateId = -1, int DepartmentId = -1, double ResourceRate = -1, double OverheadRate = -1)
        {
            try
            {
                WebApiLoginCheck("DerManagement", "update");

                #region //Request
                derManagementDA = new DerManagementDA();
                dataRequest = derManagementDA.UpdateDepartmentRate(DepartmentRateId, DepartmentId, ResourceRate, OverheadRate);
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

        #region //Delete
        #region //DeleteDepartmentRate -- 刪除部門費用率資料
        [HttpPost]
        public void DeleteDepartmentRate(int DepartmentRateId = -1)
        {
            try
            {
                WebApiLoginCheck("DerManagement", "delete");

                #region //Request
                derManagementDA = new DerManagementDA();
                dataRequest = derManagementDA.DeleteDepartmentRate(DepartmentRateId);
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