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
    public class QmsBasicInformationController : WebController
    {
        private QmsBasicInformationDA qmsBasicInformationDA = new QmsBasicInformationDA();

        #region //View
        public ActionResult DefectManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RepairManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RepairGroup()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RepairClass()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult RepairCause()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QcItemManagement()
        {
            ViewLoginCheck();

            return View();
        }
        public ActionResult QcItemManagementNew()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QcMachineModeManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QcItemPrincipleManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QcItemCodingManagement()
        {
            ViewLoginCheck();

            return View();
        }

        public ActionResult QcTheoryWorkTime()
        {
            ViewLoginCheck();
            return View();
        }

        #endregion

        #region //Get
        #region //GetDefect 取得異常狀態資料 -- Ann 2022-06-14
        [HttpPost]
        public void GetDefect()
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetDefect();
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

        #region //GetNgGroup 取得異常群組 -- Ann 2022-06-15
        [HttpPost]
        public void GetNgGroup(int GroupId = -1, string Status = "A")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetNgGroup(GroupId, Status);
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

        #region //GetNgClass 取得異常類別 -- Ann 2022-06-15
        [HttpPost]
        public void GetNgClass(int GroupId = -1, int ClassId = -1, string Status = "A")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetNgClass(GroupId, ClassId, Status);
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

        #region //GetNgCause 取得異常原因 -- Ann 2022-10-11
        [HttpPost]
        public void GetNgCause(int CauseId = -1, int ClassId = -1, string CauseNo = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetNgCause(CauseId, ClassId, CauseNo, Status);
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

        #region //GetRepair 取得維修狀態資料 -- Ann 2022-06-21
        [HttpPost]
        public void GetRepair()
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetRepair();
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

        #region //GetRepairGroup 取得維修群組資料 -- Ann 2022-06-21
        [HttpPost]
        public void GetRepairGroup(string GroupNo = "", string GroupName = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetRepairGroup(GroupNo, GroupName);
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

        #region //GetRepairClass 取得維修類別資料 -- Ann 2022-06-21
        [HttpPost]
        public void GetRepairClass(int GroupId = -1, string ClassNo = "", string ClassName = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetRepairClass(GroupId, ClassNo, ClassName);
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

        #region //GetRepairCause 取得維修原因資料 -- Ann 2022-06-21
        [HttpPost]
        public void GetRepairCause(int ClassId = -1, int CauseId = -1, string CauseNo = "", string CauseName = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetRepairCause(ClassId, CauseId, CauseNo, CauseName);
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

        #region //GetQcItemAll 取得量測項目 -- Ted 2022-10-03
        [HttpPost]
        public void GetQcItemAll()
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcItemAll();
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

        #region //GetQcGroup 取得量測群組資料 -- Ted 2022-10-03
        [HttpPost]
        public void GetQcGroup(string QcGroupNo = "", string QcGroupName = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcGroup(QcGroupNo, QcGroupName);
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

        #region //GetQcClass 取得量測類別資料 -- Ted 2022-10-03
        [HttpPost]
        public void GetQcClass(int QcGroupId = -1, string QcClassNo = "", string QcClassName = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcClass(QcGroupId, QcClassNo, QcClassName);
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

        //#region //GetQcItem 取得量測項目資料 -- Ted 2022-10-03
        //[HttpPost]
        //public void GetQcItem(int QcClassId = -1, string QcItemNo = "", string QcItemName = "")
        //{
        //    try
        //    {
        //        WebApiLoginCheck("QcItemManagement", "read");

        //        #region //Request
        //        dataRequest = qmsBasicInformationDA.GetQcItem(QcClassId, QcItemNo, QcItemName);
        //        #endregion

        //        #region //Response
        //        jsonResponse = BaseHelper.DAResponse(dataRequest);
        //        #endregion
        //    }
        //    catch (Exception e)
        //    {
        //        #region //Response
        //        jsonResponse = JObject.FromObject(new
        //        {
        //            status = "error",
        //            msg = e.Message
        //        });
        //        #endregion

        //        logger.Error(e.Message);
        //    }

        //    Response.Write(jsonResponse.ToString());
        //}
        //#endregion

        #region //GetQcItem 取得量測項目
        [HttpPost]
        public void GetQcItem(int QcItemId = -1, string QcItemNo = "",int QcProdId =-1, string QcType = "", string Status = "", string Remark = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcItem(QcItemId, QcItemNo, QcProdId, QcType, Status, Remark
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

        #region //GetQiDetail 取得量測項目機型單身
        [HttpPost]
        public void GetQiDetail(int QcItemId = -1, int QcMachineModeId = -1, string QcMachineModeNo =""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "detail,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQiDetail(QcItemId, QcMachineModeId, QcMachineModeNo
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

        #region //GetQcMachineMode 取得量測機型
        [HttpPost]
        public void GetQcMachineMode(int QcMachineModeId = -1, string QcMachineModeNo = "", string QcMachineModeName = "", string QcMachineModeNumber = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcMachineMode(QcMachineModeId, QcMachineModeNo, QcMachineModeName, QcMachineModeNumber
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

        #region //GetQmmDetail 取得量測機型 - 機台單身
        [HttpPost]
        public void GetQmmDetail(int QmmDetailId = -1, int QcMachineModeId = -1 ,string MachineNo = "", string MachineName ="", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "detail,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQmmDetail(QmmDetailId, QcMachineModeId, MachineNo, MachineName, Status
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

        #region //GetQcGroupNew 取得量測群組資料(新)
        [HttpPost]
        public void GetQcGroupNew(int QcGroupId = -1, string QcGroupNo = "", string QcGroupName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcGroupNew(QcGroupId, QcGroupNo, QcGroupName, Status
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

        #region //GetQcClassNew 取得量測類別資料(新)
        [HttpPost]
        public void GetQcClassNew(int QcClassId = -1, int QcGroupId = -1, string QcClassNo = "", string QcClassName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcClassNew(QcClassId, QcGroupId, QcClassNo, QcClassName, Status
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

        #region //GetQcItemNew 取得量測項目資料(新)
        [HttpPost]
        public void GetQcItemNew(int QcItemId = -1, int QcClassId = -1, int QcGroupId = -1, string QcQcItemNo = "", string QcQcItemName = "", string QicNo =""
            , string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcItemNew(QcItemId, QcClassId, QcGroupId, QcQcItemNo, QcQcItemName, QicNo, Status
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

        #region //GetQcClassNoMax 取得量測類別最大號 
        [HttpPost]
        public void GetQcClassNoMax(int QcGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "read");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcClassNoMax(QcGroupId);
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

        #region //GetQcItemPrinciple 取得量測項目編碼原則
        [HttpPost]
        public void GetQcItemPrinciple(int PrincipleId = -1, int QcClassId = -1, int QmmDetailId = -1, string PrincipleNo = "", string PrincipleDesc = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemPrincipleManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcItemPrinciple(PrincipleId, QcClassId, QmmDetailId, PrincipleNo, PrincipleDesc
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

        #region //GetPrincipleDetail 取得量測項目編碼原則 附加欄位
        [HttpPost]
        public void GetPrincipleDetail(int PdId = -1, int PrincipleId = -1
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemPrincipleManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetPrincipleDetail(PdId, PrincipleId
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

        #region //GetQcItemCoding 取得項目編碼規則管理
        [HttpPost]
        public void GetQcItemCoding(int QicId = -1, string QicNo = "", string QicName = "", string Status = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemCodingManagement", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcItemCoding(QicId, QicNo, QicName, Status
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

        #region //GetLotNumberByQc 取得批號資料
        [HttpPost]
        public void GetLotNumberByQc(int MoId = -1, int MoProcessId = -1, string ItemValue = "")
        {
            try
            {
                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetLotNumberByQc(MoId, MoProcessId, ItemValue);
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

        #region //GetBarcodeAttributeByQc 取得條碼屬性資料
        [HttpPost]
        public void GetBarcodeAttributeByQc(int MoId = -1, int MoProcessId = -1, string ItemNo = "")
        {
            try
            {
                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetBarcodeAttributeByQc(MoId, MoProcessId, ItemNo);
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

        #region //GetQcItemByQc 取得量測項目資料
        [HttpPost]
        public void GetQcItemByQc()
        {
            try
            {
                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcItemByQc();
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

        #region //GetQcMeasureInTheoryWorkTime -- 取得量測項目預測工時列表
        [HttpPost]
        public void GetQcMeasureInTheoryWorkTime(int QmwtId = -1, string ProductType  = "", int QmmDetailId = -1, string QicNo = ""/*, decimal WorkTime*/
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcTheoryWorkTime", "read,constrained-data");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.GetQcMeasureInTheoryWorkTime(QmwtId, ProductType, QmmDetailId, QicNo, OrderBy, PageIndex, PageSize);
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
        #region //AddNgGroup 新增異常群組 -- Ann 2022-06-14
        [HttpPost]
        public void AddNgGroup(string GroupNo = "", string GroupName = "", string Status = "", string GroupDesc = "")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "add-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddNgGroup(GroupNo, GroupName, Status, GroupDesc);
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

        #region //AddNgClass 新增異常類別 -- Ann 2022-06-15
        [HttpPost]
        public void AddNgClass(int GroupId = -1, string ClassNo = "", string ClassName = "", string ClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "add-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddNgClass(GroupId, ClassNo, ClassName, ClassDesc, Status);
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

        #region //AddNgCause 新增異常原因 -- Ann 2022-06-15
        [HttpPost]
        public void AddNgCause(int ClassId = -1, string CauseNo = "", string CauseName = "", string CauseDesc = ""
            , int ResponsibleDepartment = -1, string Status = "A")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "add-cause");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddNgCause(ClassId, CauseNo, CauseName, CauseDesc
                    , ResponsibleDepartment, Status);
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

        #region //AddRepairGroup 新增維修群組資料 -- Ann 2022-06-21
        [HttpPost]
        public void AddRepairGroup(string GroupNo = "", string GroupName = "", string GroupDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "add-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddRepairGroup(GroupNo, GroupName, GroupDesc, Status);
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

        #region //AddRepairClass 新增維修類別資料 -- Ann 2022-06-21
        [HttpPost]
        public void AddRepairClass(int GroupId = -1, string ClassNo = "", string ClassName = "", string ClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "add-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddRepairClass(GroupId, ClassNo, ClassName, ClassDesc, Status);
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

        #region //AddRepairCause 新增維修原因資料 -- Ann 2022-06-21
        [HttpPost]
        public void AddRepairCause(int ClassId = -1, string CauseNo = "", string CauseName = "", string CauseDesc = ""
            ,int ResponsibleDepartment = -1, string Status = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "add-cause");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddRepairCause(ClassId, CauseNo, CauseName, CauseDesc
                    , ResponsibleDepartment, Status);
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

        #region //AddQcGroup 新增量測群組資料 -- Ted 2022-10-03
        [HttpPost]
        public void AddQcGroup(string QcGroupNo = "", string QcGroupName = "", string QcGroupDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "add-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcGroup(QcGroupNo, QcGroupName, QcGroupDesc, Status);
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

        #region //AddQcClass 新增量測類別資料 -- Ted 2022-10-03
        [HttpPost]
        public void AddQcClass(int QcGroupId = -1, string QcClassNo = "", string QcClassName = "", string QcClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "add-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcClass(QcGroupId, QcClassNo, QcClassName, QcClassDesc, Status);
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

        #region //AddQcItem 新增量測項目 -- Ted 2022-10-03
        [HttpPost]
        public void AddQcItem(int QcClassId = -1, string QcItemNo = "", string QcItemName = "", string QcItemDesc = ""
            , string QcType = "", string QcItemType = "", int QcProdId =-1, string Remark = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "add-item");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcItem(QcClassId, QcItemNo, QcItemName, QcItemDesc
                    , QcType, QcItemType , QcProdId, Remark, Status);
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

        #region //AddQcItemBatch 新增量測項目(批量) -- Shintokuro 2024-05-17
        [HttpPost]
        public void AddQcItemBatch(string BatchFormData = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "add-item");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcItemBatch(BatchFormData);
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

        #region //AddQiDetail 新增量測項目機型單身
        [HttpPost]
        public void AddQiDetail(int QcItemId = -1, int QcMachineModeId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "detail");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQiDetail(QcItemId, QcMachineModeId);
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

        #region //AddQcMachineMode 新增量測機型
        [HttpPost]
        public void AddQcMachineMode(string QcMachineModeNo = "", string QcMachineModeName = "", string QcMachineModeDesc = "", string ItemNo = "")
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "add");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcMachineMode(QcMachineModeNo, QcMachineModeName, QcMachineModeDesc, ItemNo);
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

        #region //AddQiMachineMode 取得量測機型 - 機台單身
        [HttpPost]
        public void AddQiMachineMode(int QcMachineModeId = -1, string MachineId = "")
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "detail");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQiMachineMode(QcMachineModeId, MachineId);
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

        #region //AddQcItemPrinciple 新增量測項目編碼原則
        [HttpPost]
        public void AddQcItemPrinciple(int QcClassId = -1, int QmmDetailId = -1, string PrincipleNo = "", string PrincipleDesc = "")
        {
            try
            {
                WebApiLoginCheck("QcItemPrincipleManagement", "add");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcItemPrinciple(QcClassId, QmmDetailId, PrincipleNo, PrincipleDesc);
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

        #region //AddQcItemCoding 新增項目編碼規則管理
        [HttpPost]
        public void AddQcItemCoding(string QicNo = "", string QicName = "", string QicDesc = "")
        {
            try
            {
                WebApiLoginCheck("QcItemCodingManagement", "add");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcItemCoding(QicNo, QicName, QicDesc);
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

        #region //AddQcMeasureInTheoryWorkTime -- 新增量測項目預測工時
        [HttpPost]
        public void AddQcMeasureInTheoryWorkTime(int QmmDetailId = -1, string ProductType = "", int QicId  =-1, string MeasureSize  ="", float WorkTime = -1)
        {
            try
            {
                WebApiLoginCheck("QcTheoryWorkTime", "add");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.AddQcMeasureInTheoryWorkTime(QmmDetailId, ProductType, QicId, MeasureSize, WorkTime);
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
        #region //UpdateNgGroup 更新異常群組 -- Ann --2022.06.15
        [HttpPost]
        public void UpdateNgGroup(int GroupId = -1, string GroupNo = "", string GroupName = "", string GroupDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "update-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateNgGroup(GroupId, GroupNo, GroupName, GroupDesc, Status);
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

        #region //UpdateNgClass 更新異常類別 -- Ann --2022.06.15
        [HttpPost]
        public void UpdateNgClass(int ClassId = -1, int GroupId = -1, string ClassNo = "", string ClassName = "", string ClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "update-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateNgClass(ClassId, GroupId, ClassNo, ClassName, ClassDesc, Status);
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

        #region //UpdateNgCause 更新異常原因 -- Ann --2022.06.15
        [HttpPost]
        public void UpdateNgCause(int CauseId = -1, int ClassId = -1, string CauseNo = "", string CauseName = "-"
            ,string DepartmentNo = "", string CauseDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "update-cause");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateNgCause(CauseId, ClassId, CauseNo, CauseName
                    , DepartmentNo, CauseDesc, Status);
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

        #region //UpdateRepairGroup 更新維修群組資料 -- Ann --2022.06.21
        [HttpPost]
        public void UpdateRepairGroup(int GroupId = -1, string GroupNo = "", string GroupName = "", string GroupDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "update-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateRepairGroup(GroupId, GroupNo, GroupName, GroupDesc, Status);
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

        #region //UpdateRepairClass 更新維修類別資料 -- Ann --2022.06.21
        [HttpPost]
        public void UpdateRepairClass(int ClassId = -1, int GroupId = -1, string ClassNo = "", string ClassName = "", string ClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "update-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateRepairClass(ClassId, GroupId, ClassNo, ClassName, ClassDesc, Status);
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

        #region //UpdateNgCause 更新維修原因資料 -- Ann --2022.06.21
        [HttpPost]
        public void UpdateRepairCause(int CauseId = -1, int ClassId = -1, string CauseNo = "", string CauseName = "-"
            , string DepartmentNo = "", string CauseDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "update-cause");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateRepairCause(CauseId, ClassId, CauseNo, CauseName
                    , DepartmentNo, CauseDesc, Status);
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

        #region //UpdateQcGroup 更新量測群組資料 -- Ted --2022.10.03
        [HttpPost]
        public void UpdateQcGroup(int QcGroupId = -1, string QcGroupNo = "", string QcGroupName = "", string QcGroupDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "update-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcGroup(QcGroupId, QcGroupNo, QcGroupName, QcGroupDesc, Status);
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

        #region //UpdateQcClass 更新量測類別資料 -- Ted --2022.10.03
        [HttpPost]
        public void UpdateQcClass(int QcClassId = -1, int QcGroupId = -1, string QcClassNo = "", string QcClassName = "", string QcClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "update-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcClass(QcClassId, QcGroupId, QcClassNo, QcClassName, QcClassDesc, Status);
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

        #region //UpdateQcItem 更新量測項目資料 -- Ted --2022.10.03
        [HttpPost]
        public void UpdateQcItem(int QcItemId = -1, int QcClassId = -1, string QcItemNo = "", string QcItemName = "-", string QcItemDesc = ""
            , string QcType = "", string QcItemType = "", string Remark = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "update-item");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcItem(QcItemId, QcClassId, QcItemNo, QcItemName
                    , QcItemDesc, QcType, QcItemType, Remark, Status);
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

        #region //UpdateQiDetail 更新量測項目機型單身
        [HttpPost]
        public void UpdateQiDetail(int QiDetailId = -1, int QcItemId = -1, int QcMachineModeId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "detail");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQiDetail(QiDetailId, QcItemId, QcMachineModeId);
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

        #region //UpdateQcMachineMode 更新量測機型資料
        [HttpPost]
        public void UpdateQcMachineMode(int QcMachineModeId = -1, string QcMachineModeNo = "", string QcMachineModeName = "", string QcMachineModeDesc = ""
            ,string ItemNo = "")
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "update");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcMachineMode(QcMachineModeId, QcMachineModeNo, QcMachineModeName, QcMachineModeDesc
                    , ItemNo);
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

        #region //UpdateQmmDetailStatus 更新量測機型的機台單身狀態
        [HttpPost]
        public void UpdateQmmDetailStatus(int QmmDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "detail");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQmmDetailStatus(QmmDetailId);
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

        #region //UpdateQcGroupNew 更新量測群組資料(新) 
        [HttpPost]
        public void UpdateQcGroupNew(int QcGroupId = -1, string QcGroupName = "", string QcGroupDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "update-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcGroupNew(QcGroupId, QcGroupName, QcGroupDesc, Status);
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

        #region //UpdateQcClassNew 更新量測類別資料(新)
        [HttpPost]
        public void UpdateQcClassNew(int QcClassId = -1, int QcGroupId = -1, string QcClassName = "", string QcClassDesc = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "update-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcClassNew(QcClassId, QcGroupId, QcClassName, QcClassDesc, Status);
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

        #region //UpdateQcItemNew 更新量測項目資料(新)
        [HttpPost]
        public void UpdateQcItemNew(int QcItemId = -1, string QcItemName = "-", string QcItemDesc = ""
            , string QcType = "", string QcItemType = "", string Remark = "", string Status = "")
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "update-item");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcItemNew(QcItemId, QcItemName
                    , QcItemDesc, QcType, QcItemType, Remark, Status);
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

        #region //UpdateQcItemPrinciple 更新量測項目編碼原則
        [HttpPost]
        public void UpdateQcItemPrinciple(int PrincipleId = -1, int QcClassId = -1, int QmmDetailId = -1, string PrincipleNo = "", string PrincipleDesc = "")
        {
            try
            {
                WebApiLoginCheck("QcItemPrincipleManagement", "update");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcItemPrinciple(PrincipleId, QcClassId, QmmDetailId, PrincipleNo, PrincipleDesc);
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

        #region //UpdateQcItemCoding 更新項目編碼規則管理
        [HttpPost]
        public void UpdateQcItemCoding(int QicId, string QicName, string QicDesc)
        {
            try
            {
                WebApiLoginCheck("QcItemCodingManagement", "update");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcItemCoding(QicId, QicName, QicDesc);
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

        #region //UpdateQcItemCodingStatus 更新項目編碼規則管理狀態
        [HttpPost]
        public void UpdateQcItemCodingStatus(int QicId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemCodingManagement", "status-switch");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcItemCodingStatus(QicId);
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

        #region //UpdateQcMeasureInTheoryWorkTime -- 編輯量測項目預測工時
        [HttpPost]
        public void UpdateQcMeasureInTheoryWorkTime(int QmwtId, int QmmDetailId, string ProductType, int QicId, string MeasureSize, float WorkTime)
        {
            try
            {
                WebApiLoginCheck("QcTheoryWorkTime", "update");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.UpdateQcMeasureInTheoryWorkTime(QmwtId, QmmDetailId, ProductType, QicId, MeasureSize, WorkTime);
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
        #region //DeleteNgGroup 刪除異常群組 -- Ann --2022.06.16
        [HttpPost]
        public void DeleteNgGroup(int GroupId = -1)
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "delete-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteNgGroup(GroupId);
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

        #region //DeleteNgClass 刪除異常類別 -- Ann --2022.06.16
        [HttpPost]
        public void DeleteNgClass(int ClassId = -1)
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "delete-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteNgClass(ClassId);
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

        #region //DeleteNgCause 刪除異常原因 -- Ann --2022.06.16
        [HttpPost]
        public void DeleteNgCause(int CauseId = -1)
        {
            try
            {
                WebApiLoginCheck("DefectManagement", "delete-cause");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteNgCause(CauseId);
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

        #region //DeleteRepairGroup 刪除維修群組資料 -- Ann --2022.06.21
        [HttpPost]
        public void DeleteRepairGroup(int GroupId = -1)
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "delete-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteRepairGroup(GroupId);
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

        #region //DeleteRepairClass 刪除維修類別資料 -- Ann --2022.06.21
        [HttpPost]
        public void DeleteRepairClass(int ClassId = -1)
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "delete-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteRepairClass(ClassId);
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

        #region //DeleteRepairCause 刪除維修原因資料 -- Ann --2022.06.21
        [HttpPost]
        public void DeleteRepairCause(int CauseId = -1)
        {
            try
            {
                WebApiLoginCheck("RepairManagement", "delete-cause");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteRepairCause(CauseId);
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

        #region //DeleteQcGroup 刪除量測群組資料 -- Ted --2022.10.03
        [HttpPost]
        public void DeleteQcGroup(int QcGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "delete-group");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQcGroup(QcGroupId);
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

        #region //DeleteQcClass 刪除量測類別資料 -- Ted --2022.10.03
        [HttpPost]
        public void DeleteQcClass(int QcClassId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "delete-class");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQcClass(QcClassId);
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

        #region //DeleteQcItem 刪除量測項目資料 -- Ted --2022.10.03
        [HttpPost]
        public void DeleteQcItem(int QcItemId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "delete-item");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQcItem(QcItemId);
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

        #region //DeleteQiDetail 刪除量測項目機型單身
        [HttpPost]
        public void DeleteQiDetail(int QiDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemManagementNew", "detail");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQiDetail(QiDetailId);
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

        #region //DeleteQiMachineMode 刪除量測機型
        [HttpPost]
        public void DeleteQiMachineMode(int QcMachineModeId = -1)
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "delete");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQiMachineMode(QcMachineModeId);
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

        #region //DeleteQmmDetail 刪除量測機型的機台單身
        [HttpPost]
        public void DeleteQmmDetail(int QmmDetailId = -1)
        {
            try
            {
                WebApiLoginCheck("QcMachineModeManagement", "detail");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQmmDetail(QmmDetailId);
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

        #region //DeleteQcItemPrinciple 刪除量測項目編碼原則
        [HttpPost]
        public void DeleteQcItemPrinciple(int PrincipleId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemPrincipleManagement", "delete");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQcItemPrinciple(PrincipleId);
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

        #region //DeleteQcItemCoding 刪除項目編碼規則管理
        [HttpPost]
        public void DeleteQcItemCoding(int QicId = -1)
        {
            try
            {
                WebApiLoginCheck("QcItemCodingManagement", "delete");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQcItemCoding(QicId);
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

        #region //DeleteQcMeasureInTheoryWorkTime 刪除量測項目預測工時
        [HttpPost]
        public void DeleteQcMeasureInTheoryWorkTime(int QmwtId = -1)
        {
            try
            {
                WebApiLoginCheck("QcTheoryWorkTime", "delete");

                #region //Request
                qmsBasicInformationDA = new QmsBasicInformationDA();
                dataRequest = qmsBasicInformationDA.DeleteQcMeasureInTheoryWorkTime(QmwtId);
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