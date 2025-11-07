using Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using MESDA;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Data.SqlClient;

using Org.BouncyCastle.Asn1.Ocsp;
using System.Drawing;
using System.Net.Sockets;
using System.Threading;
using ZXing;
using ZXing.Common;

using System.Security.Cryptography;
using MAMODA;

namespace Business_Manager.Controllers
{
    public class MamoBasicInformationController : WebController
    {
        private MamoBasicInformationDA mamoBasicInformationDA = new MamoBasicInformationDA();

        #region //View

        // GET: MamoBasicInformation
        public ActionResult Team()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult Channel()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion //View

        #region //Get

        #region //GetMamoTeams                  取得團隊資料
        [HttpPost]
        public void GetMamoTeams(
            string  TeamId = "",
            string  MamoTeamId = "",
            string  TeamName = "",
            string  Remark = "",
            string  Status = "",
            string  OrderBy = "",
            int     PageIndex = -1,
            int     PageSize = -1
        )
        {
            try
            {
                WebApiLoginCheck("Team", "read,constrained-data");

                #region //Request
                mamoBasicInformationDA = new MamoBasicInformationDA();
                dataRequest = mamoBasicInformationDA.GetMamoTeams(TeamId, MamoTeamId, TeamName, Remark, Status, OrderBy, PageIndex, PageSize);

                #endregion //Request

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

        #endregion //GetMamoTeams   取得團隊資料

        #region //GetMamoTeamMembers            取得團隊成員資料
        [HttpPost]
        public void GetMamoTeamMembers(
            int MemberId = -1,
            int TeamId = -1,
            int CompanyId = -1,
            string Departments = "",
            string UserNo = "",
            string UserName = "",
            string OrderBy = "",
            int PageIndex = -1,
            int PageSize = -1
        )
        {
            try
            {
                WebApiLoginCheck("Team", "allow-user");

                #region //Request
                mamoBasicInformationDA = new MamoBasicInformationDA();
                dataRequest = mamoBasicInformationDA.GetMamoTeamMembers(MemberId, TeamId, CompanyId
                    , Departments, UserNo, UserName
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
        #endregion //GetMamoTeamMembers 取得團隊成員資料

        #region //GetMamoChannels               取得頻道資訊 TODO
        [HttpPost]
        public void GetMamoChannels(
            string  ChannelId = "",
            string  MamoChannelId = "",
            string  TeamId = "",
            string  ChannelName = "",
            string  ChannelNo = "",
            string  Status = "",
            string  OrderBy = "",
            int     PageIndex = -1,
            int     PageSize = -1
        )
        {
            try
            {
                WebApiLoginCheck("Channel", "read,constrained-data");

                #region //Request
                mamoBasicInformationDA = new MamoBasicInformationDA();
                dataRequest = mamoBasicInformationDA.GetMamoChannels(ChannelId, MamoChannelId, TeamId, ChannelName, ChannelNo, Status, OrderBy, PageIndex, PageSize);

                #endregion //Request

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


        #endregion //GetMamoChannels(TeamId)    取得頻道資訊

        #region //GetMamoChannelMembers         取得頻道團員資訊 TODO

        #endregion //GetMamoChannelMembers(ChannelId)   取得頻道團員資訊

        #endregion //Get

        #region //Add

        #region //AddMamoTeams                  建立團隊
        public void AddMamoTeams(
            string TeamName,
            string Remark
        )
        {
            try
            {
                int UserId = -1;
                int CompanyId = -1;
                WebApiLoginCheck("Team", "add");
                mamoBasicInformationDA = new MamoBasicInformationDA();

                #region //GetUserDetail
                UserId = Convert.ToInt32(Session["UserId"]); ;
                dataRequest = mamoBasicInformationDA.GetUserDetail(UserId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string UserNo = jsonResponse["result"][0]["UserNo"].ToString();

                #endregion

                #region //GetCompanyDetail
                CompanyId = Convert.ToInt32(Session["UserCompany"]);
                dataRequest = mamoBasicInformationDA.GetCompanyDetail(CompanyId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string CompanyNo = jsonResponse["result"][0]["CompanyNo"].ToString();

                #endregion

                #region //request
                dataRequest = mamoBasicInformationDA.AddMamoTeams(CompanyNo, UserId, TeamName, Remark);

                #endregion //request

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                #endregion //Response
            }

            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion //Response

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion //AddMamoTeams   建立團隊

        #region //AddMamoTeamMembers            新增團隊成員
        [HttpPost]
        public void AddMamoTeamMembers(
            int TeamId = -1,
            string Users = ""
        )
        {
            try
            {
                WebApiLoginCheck("Team", "allow-user");

                int UserId = -1;
                int CompanyId = -1;
                WebApiLoginCheck("Team", "add");
                mamoBasicInformationDA = new MamoBasicInformationDA();

                #region //GetUserDetail
                UserId = Convert.ToInt32(Session["UserId"]); ;
                dataRequest = mamoBasicInformationDA.GetUserDetail(UserId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string UserNo = jsonResponse["result"][0]["UserNo"].ToString();

                #endregion //GetUserDetail

                #region //GetCompanyDetail
                CompanyId = Convert.ToInt32(Session["UserCompany"]);
                dataRequest = mamoBasicInformationDA.GetCompanyDetail(CompanyId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string CompanyNo = jsonResponse["result"][0]["CompanyNo"].ToString();

                #endregion //GetCompanyDetail

                #region //Request
                dataRequest = mamoBasicInformationDA.AddMamoTeamMembers(CompanyNo, UserId, TeamId, Users);

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

                #endregion //Response

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }

        #endregion //AddMamoTeamMembers 新增團隊成員

        #region //AddMamoChannels               建立頻道 TODO
        public void AddMamoChannels(
            string ChannelName,
            int TeamId,
            string Remark
        )
        {
            try
            {
                int UserId = -1;
                int CompanyId = -1;
                WebApiLoginCheck("Team", "add");
                mamoBasicInformationDA = new MamoBasicInformationDA();

                #region //GetUserDetail
                UserId = Convert.ToInt32(Session["UserId"]); ;
                dataRequest = mamoBasicInformationDA.GetUserDetail(UserId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string UserNo = jsonResponse["result"][0]["UserNo"].ToString();

                #endregion

                #region //GetCompanyDetail
                CompanyId = Convert.ToInt32(Session["UserCompany"]);
                dataRequest = mamoBasicInformationDA.GetCompanyDetail(CompanyId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string CompanyNo = jsonResponse["result"][0]["CompanyNo"].ToString();

                #endregion

                #region //request
                dataRequest = mamoBasicInformationDA.AddMamoChannels(CompanyNo, UserId, ChannelName, TeamId, Remark);

                #endregion //request

                #region //Response
                jsonResponse = BaseHelper.DAResponse(dataRequest);

                #endregion //Response
            }

            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion //Response

                logger.Error(e.Message);
            }

            Response.Write(jsonResponse.ToString());
        }


        #endregion //AddMamoChannels    建立頻道

        #region //AddMamoChannelMembers         新增頻道成員 TODO

        #endregion //AddMamoChannelMembers  新增頻道成員

        #endregion //Add

        #region //Update

        #region //UpdateMamoTeamsStatus         團隊狀態更新
        [HttpPost]
        public void UpdateMamoTeamsStatus(int TeamId = -1)
        {
            try
            {
                int UserId = -1;
                int CompanyId = -1;
                WebApiLoginCheck("Team", "status-switch");
                mamoBasicInformationDA = new MamoBasicInformationDA();

                #region //GetUserDetail
                UserId = Convert.ToInt32(Session["UserId"]); ;
                dataRequest = mamoBasicInformationDA.GetUserDetail(UserId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string UserNo = jsonResponse["result"][0]["UserNo"].ToString();

                #endregion

                #region //GetCompanyDetail
                CompanyId = Convert.ToInt32(Session["UserCompany"]);
                dataRequest = mamoBasicInformationDA.GetCompanyDetail(CompanyId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string CompanyNo = jsonResponse["result"][0]["CompanyNo"].ToString();

                #endregion

                #region //Request
                dataRequest = mamoBasicInformationDA.UpdateMamoTeamsStatus(CompanyNo, UserId, TeamId);

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

        #endregion //UpdateMamoTeamsStatus  團隊狀態更新

        #region //UpdateMamoChannelsStatus      頻道狀態更新 TODO
        #endregion //UpdateMamoChannelsStatus   頻道狀態更新

        #region //UpdateMamoTeamsInfo           團隊名稱更新 TODO
        #endregion //UpdateMamoTeamsInfo        團隊名稱更新

        #region //UpdateMamoChannelsInfo        頻道名稱更新 TODO
        #endregion //UpdateMamoChannelsInfo     頻道名稱更新

        #endregion //Update

        #region //Delete

        #region //DeleteMamoTeamMembers         團隊成員資料刪除
        [HttpPost]
        public void DeleteMamoTeamMembers(
            int MemberId = -1
        )
        {
            try
            {
                WebApiLoginCheck("Team", "allow-user");

                int UserId = -1;
                int CompanyId = -1;
                WebApiLoginCheck("Team", "add");
                mamoBasicInformationDA = new MamoBasicInformationDA();

                #region //GetUserDetail
                UserId = Convert.ToInt32(Session["UserId"]); ;
                dataRequest = mamoBasicInformationDA.GetUserDetail(UserId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string UserNo = jsonResponse["result"][0]["UserNo"].ToString();

                #endregion

                #region //GetCompanyDetail
                CompanyId = Convert.ToInt32(Session["UserCompany"]);
                dataRequest = mamoBasicInformationDA.GetCompanyDetail(CompanyId);
                jsonResponse = BaseHelper.DAResponse(dataRequest);
                string CompanyNo = jsonResponse["result"][0]["CompanyNo"].ToString();

                #endregion

                #region //Request
                mamoBasicInformationDA = new MamoBasicInformationDA();
                dataRequest = mamoBasicInformationDA.DeleteMamoTeamMembers(CompanyNo, UserId, MemberId);

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
        #endregion //DeleteMamoTeamMembers      團隊成員資料刪除

        #region //DeleteMamoChannelMembers      頻道成員資料刪除 TODO

        #endregion //DeleteMamoChannelMembers   頻道成員資料刪除

        #endregion //Delete

    }
}