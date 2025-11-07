using Helpers;
using Newtonsoft.Json.Linq;
using QMSDA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Business_Manager.Controllers
{
    public class AbnormalqualityController : WebController
    {
        private AbnormalqualityDA abnormalqualityDA = new AbnormalqualityDA();

        #region //View
        public ActionResult AqReturnManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AqCountermeasureManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AqcConfirmManagment()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult AqJudgmentManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AqCountersignManagment()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult AqViewManagment()
        {
            ViewLoginCheck();

            return View();
        }
        #endregion

        #region //Get

        #region //GetAbnormalquality 取得品異單單頭
        [HttpPost]
        public void GetAbnormalquality(int AbnormalqualityId = -1, string AbnormalqualityNo = "", string CreateDate = "", string AbnormalqualityStatus = ""
            , string StartCreateDate = "" , string EndCreateDate = "", string BarcodeNo = "", string ErpNo="",int MoId = -1, int ResponsibleDeptId = -1
            , string AqBarcodeStatus = "", string JudgeConfirm = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "read");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetAbnormalquality(AbnormalqualityId, AbnormalqualityNo, CreateDate, AbnormalqualityStatus
                    , StartCreateDate, EndCreateDate, BarcodeNo, ErpNo, MoId, ResponsibleDeptId
                    , AqBarcodeStatus, JudgeConfirm
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

        #region //GetAqDetail 取得品異條碼單身
        [HttpPost]
        public void GetAqDetail(int AqBarcodeId = -1, int AbnormalqualityId = -1, string AqBarcodeStatus = "", string JudgeConfirm = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "read");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetAqDetail(AqBarcodeId, AbnormalqualityId, AqBarcodeStatus, JudgeConfirm
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

        #region //GetAqQcItem 取得品異條碼異常項目原因
        [HttpPost]
        public void GetAqQcItem(int AqBarcodeId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetAqQcItem(AqBarcodeId
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

        #region //GetAqFile 取得品異條碼原因佐證檔案
        [HttpPost]
        public void GetAqFile(int AqFileId = -1, int AqBarcodeId = -1, string AqFileStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetAqFile(AqFileId, AqBarcodeId, AqFileStatus
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

        #region //GetAqFileShow 取得品異條碼原因佐證檔案(顯示)
        [HttpPost]
        public void GetAqFileShow(int AqBarcodeId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetAqFileShow(AqBarcodeId
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

        #region //GetJudgeReturnMoProcess 取得判定退回站別清單
        [HttpPost]
        public void GetJudgeReturnMoProcess(int MoId
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetJudgeReturnMoProcess(MoId
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

        #region //GetBarcodeList 取得條碼列表
        [HttpPost]
        public void GetBarcodeList(int MoId = -1, int MoProcessId =-1, string BarcodeListView =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetBarcodeList(MoId,MoProcessId, BarcodeListView
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

        #region //GetBarcodeLettering 取得條碼的刻號
        [HttpPost]
        public void GetBarcodeLettering(string BarcodeNo = "")
        {
            try
            {
                //WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetBarcodeLettering(BarcodeNo);
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

        #region //GetAqPhrase 取得常用片語
        [HttpPost]
        public void GetAqPhrase(string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "read");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.GetAqPhrase(OrderBy, PageIndex, PageSize);
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
        #region //AddAbnormalquality 新增品異單(單頭)
        [HttpPost]
        public void AddAbnormalquality(int? MoId = -1, int? GrDetailId = -1, int? MoProcessId = -1, string QcType= "", string DocDate = "", int ViewCompanyId = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "add");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.AddAbnormalquality(MoId, GrDetailId, MoProcessId, QcType, DocDate, ViewCompanyId);
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

        #region //AddAqDetail 新增品異單 - WEB版本
        [HttpPost]
        public void AddAqDetail(int AbnormalqualityId = -1, string BarcodeIds = "", int DefectCauseId = -1, string DefectCauseDesc = ""
            ,int ResponsibleDeptId = -1, int ResponsibleUserId = -1, int SubResponsibleDeptId = -1, int SubResponsibleUserId = -1
            , int ProgrammerUserId = -1, int ConformUserId = -1, int? RepairCauseId = -1, string RepairCauseDesc = "", int? RepairCauseUserId = -1
            , int SupplierId = -1, string ChangeJudgeFlag = "")

        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "add");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.AddAqDetail(AbnormalqualityId, BarcodeIds, DefectCauseId, DefectCauseDesc
                    , ResponsibleDeptId, ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId
                    , ProgrammerUserId, ConformUserId, RepairCauseId, RepairCauseDesc, RepairCauseUserId, SupplierId, ChangeJudgeFlag);
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

        #region //AddAbnormalqualityPadProject 新增品異單 - 平板版本(到對策確認前)
        [HttpPost]
        public void AddAbnormalqualityPadProject(string AbnormalqualityData = "", string AbnormalProjectList = "")
        {
            try
            {
                //WebApiLoginCheck("AqReturnManagment", "add");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.AddAbnormalqualityPadProject(AbnormalqualityData, AbnormalProjectList);
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

        #region //AddAbnormalqualityPad 新增品異單 - 平板版本(到對策確認前)
        [HttpPost]
        public void AddAbnormalqualityPad(string AbnormalqualityData = "")
        {
            try
            {
                //WebApiLoginCheck("AqReturnManagment", "add");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.AddAbnormalqualityPad(AbnormalqualityData);
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

        #region //AddAqFile 新增品異條碼原因佐證檔案
        [HttpPost]
        public void AddAqFile(int AqBarcodeId = -1, string FileId = "")
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.AddAqFile(AqBarcodeId, FileId);
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

        #region //UpdateAqReturnDetail 更新品異單異常回報資料
        [HttpPost]
        public void UpdateAqReturnDetail(int AbnormalqualityId = -1, int AqBarcodeId = -1, int DefectCauseId = -1, string DefectCauseDesc = "", int ResponsibleDeptId = -1
            , int ResponsibleUserId = -1, int SubResponsibleDeptId = -1, int SubResponsibleUserId = -1, int ProgrammerUserId = -1, int ConformUserId = -1
            , int? RepairCauseId = -1, string RepairCauseDesc = "", int? RepairCauseUserId = -1, int SupplierId = -1, string ChangeJudgeFlag = "")
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqReturnDetail(AbnormalqualityId, AqBarcodeId, DefectCauseId, DefectCauseDesc, ResponsibleDeptId
                    , ResponsibleUserId, SubResponsibleDeptId, SubResponsibleUserId, ProgrammerUserId, ConformUserId
                    , RepairCauseId, RepairCauseDesc, RepairCauseUserId, SupplierId, ChangeJudgeFlag);
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

        #region //UpdateAqCountermeasureDetail 更新品異單異常對策資料
        [HttpPost]
        public void UpdateAqCountermeasureDetail(int AbnormalqualityId = -1, int AqBarcodeId = -1, int RepairCauseId = -1, string RepairCauseDesc = "", int RepairCauseUserId = -1,string RepairCauseList =""
            , int ResponsibleDeptId =-1, int ResponsibleUserId =-1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqCountermeasureDetail(AbnormalqualityId, AqBarcodeId, RepairCauseId, RepairCauseDesc, RepairCauseUserId, RepairCauseList
                    ,ResponsibleDeptId, ResponsibleUserId);
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

        #region //UpdateAqCountermeasureDetailBatch 更新品異單異常對策資料(批量)
        [HttpPost]
        public void UpdateAqCountermeasureDetailBatch(int AbnormalqualityId = -1, string AqBarcodeIdList = "", int RepairCauseId = -1, string RepairCauseDesc = "", int RepairCauseUserId = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqCountermeasureDetailBatch(AbnormalqualityId, AqBarcodeIdList, RepairCauseId, RepairCauseDesc, RepairCauseUserId);
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

        #region //UpdateAqcConfirmDetail 更新品異單異常對策資料確認
        [HttpPost]
        public void UpdateAqcConfirmDetail(int AbnormalqualityId = -1, int AqBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("AqcConfirmManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqcConfirmDetail(AbnormalqualityId, AqBarcodeId);
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

        #region //UpdateAqcConfirmDetailRe 更新品異單異常對策資料確認(撤銷)
        [HttpPost]
        public void UpdateAqcConfirmDetailRe(int AbnormalqualityId = -1, int AqBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("AqcConfirmManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqcConfirmDetailRe(AbnormalqualityId, AqBarcodeId);
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

        #region //UpdateAqcConfirmDetailBatch 更新品異單異常對策資料確認(批量)
        [HttpPost]
        public void UpdateAqcConfirmDetailBatch(int AbnormalqualityId = -1, string AqBarcodeIdList = "")
        {
            try
            {
                WebApiLoginCheck("AqcConfirmManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqcConfirmDetailBatch(AbnormalqualityId, AqBarcodeIdList);
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

        #region //UpdateAqJudgmentDetail 更新品異單異常判定資料
        [HttpPost]
        public void UpdateAqJudgmentDetail(int AbnormalqualityId = -1, int AqBarcodeId = -1, string JudgeStatus ="", int JudgeReturnMoProcessId = -1
            , int JudgeReturnNextMoProcessId = -1, string JudgeDesc ="", int JudgeUserId = -1, string JudgeDate = "", int ResponsibleUserId = -1, int ResponsibleDeptId = -1
            , int ReleaseQty = -1, string UserAqPhrase = "")
        {
            try
            {
                WebApiLoginCheck("AqJudgmentManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqJudgmentDetail(AbnormalqualityId, AqBarcodeId, JudgeStatus, JudgeReturnMoProcessId
                    ,JudgeReturnNextMoProcessId , JudgeDesc ,JudgeUserId, JudgeDate, ResponsibleUserId, ResponsibleDeptId
                    , ReleaseQty, UserAqPhrase);
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

        #region //UpdateAqJudgmentDetailBatch 更新品異單異常判定資料(批量)
        [HttpPost]
        public void UpdateAqJudgmentDetailBatch(int AbnormalqualityId = -1, string AqBarcodeIdList = "", string JudgeStatus = "", int JudgeReturnMoProcessId = -1
            , int JudgeReturnNextMoProcessId = -1, string JudgeDesc = "", int JudgeUserId = -1, string JudgeDate = "", int ResponsibleDeptId = -1, int ResponsibleUserId = -1
            , int ReleaseQty = -1, string UserAqPhrase = "")
        {
            try
            {
                WebApiLoginCheck("AqJudgmentManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqJudgmentDetailBatch(AbnormalqualityId, AqBarcodeIdList, JudgeStatus, JudgeReturnMoProcessId
                    ,JudgeReturnNextMoProcessId , JudgeDesc, JudgeUserId, JudgeDate, ResponsibleDeptId, ResponsibleUserId, ReleaseQty, UserAqPhrase);
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

        #region //UpdateAqCountersignDetail 更新品異單異常判定資料會簽
        [HttpPost]
        public void UpdateAqCountersignDetail(int AbnormalqualityId = -1, int AqBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("AqcConfirmManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqCountersignDetail(AbnormalqualityId, AqBarcodeId);
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

        #region //UpdateAqCountersignDetailRe 更新品異單異常判定資料會簽(撤銷)
        [HttpPost]
        public void UpdateAqCountersignDetailRe(int AbnormalqualityId = -1, int AqBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("AqcConfirmManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqCountersignDetailRe(AbnormalqualityId, AqBarcodeId);
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

        #region //UpdateAqCountersignDetailBatch 更新品異單異常判定資料會簽(批量)
        [HttpPost]
        public void UpdateAqCountersignDetailBatch(int AbnormalqualityId = -1, string AqBarcodeIdList = "")
        {
            try
            {
                WebApiLoginCheck("AqcConfirmManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.UpdateAqCountersignDetailBatch(AbnormalqualityId, AqBarcodeIdList);
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

        #region //DeleteAbnormalquality 刪除品異單
        [HttpPost]
        public void DeleteAbnormalquality(int AbnormalqualityId = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.DeleteAbnormalquality(AbnormalqualityId);
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

        #region //DeleteAqDetail 刪除品異單單身(異常條碼) 
        [HttpPost]
        public void DeleteAqDetail(int AqBarcodeId = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.DeleteAqDetail(AqBarcodeId);
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

        #region //DeleteAqFileId 品異條碼原因佐證檔案
        [HttpPost]
        public void DeleteAqFileId(int AqFileId = -1)
        {
            try
            {
                WebApiLoginCheck("AqReturnManagment", "detail");

                #region //Request
                abnormalqualityDA = new AbnormalqualityDA();
                dataRequest = abnormalqualityDA.DeleteAqFileId(AqFileId);
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