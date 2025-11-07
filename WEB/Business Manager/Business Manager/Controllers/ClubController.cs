using EBPDA;
using Helpers;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class ClubController : WebController
    {
        private ClubDA clubDA = new ClubDA();

        #region //View
        public ActionResult ClubManagement()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get
        #region //GetClub 取得社團資料
        [HttpPost]
        public void GetClub(int ClubId = -1, string ClubName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "read");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.GetClub(ClubId, ClubName, Status
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

        #region //GetClubMember 取得社團成員資料
        [HttpPost]
        public void GetClubMember(int MemberId = -1, int ClubId = -1, string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "member");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.GetClubMember(MemberId, ClubId, SearchKey
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
        #region //AddClub 社團資料新增
        [HttpPost]
        public void AddClub(string ClubName = "", int ClubApplicant = -1, string EstablishedDate = "", string ClubProperty = ""
            , string ClubGoal = "", string Appropriation = "", string ActiveRegion = "")
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "add");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.AddClub(ClubName, ClubApplicant, EstablishedDate, ClubProperty
                    , ClubGoal, Appropriation, ActiveRegion);
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

        #region //AddClubMember 社團成員資料新增
        [HttpPost]
        public void AddClubMember(int ClubId = -1, int ClubJobId = -1, string Users = "")
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "member");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.AddClubMember(ClubId, ClubJobId, Users);
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
        #region //UpdateClub 社團資料更新
        [HttpPost]
        public void UpdateClub(int ClubId = -1, string ClubName = "", int ClubApplicant = -1, string EstablishedDate = ""
            , string ClubProperty = "", string ClubGoal = "", string Appropriation = "", string ActiveRegion = "")
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "update");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.UpdateClub(ClubId, ClubName, ClubApplicant, EstablishedDate
                    , ClubProperty, ClubGoal, Appropriation, ActiveRegion);
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

        #region //UpdateClubStatus 社團狀態更新
        [HttpPost]
        public void UpdateClubStatus(int ClubId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "status-switch");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.UpdateClubStatus(ClubId);
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
        #region //DeleteClub 社團資料刪除
        [HttpPost]
        public void DeleteClub(int ClubId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "delete");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.DeleteClub(ClubId);
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

        #region //DeleteClubMember 社團成員資料刪除
        [HttpPost]
        public void DeleteClubMember(int MemberId = -1)
        {
            try
            {
                WebApiLoginCheck("ClubManagement", "member");

                #region //Request
                clubDA = new ClubDA();
                dataRequest = clubDA.DeleteClubMember(MemberId);
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