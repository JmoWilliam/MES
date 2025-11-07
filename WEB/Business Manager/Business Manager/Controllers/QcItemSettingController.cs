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
    public class QcItemSettingController : WebController
    {
        private QcItemSettingDA qcItemSettingDA = new QcItemSettingDA();

        public ActionResult QcNotice()
        {
            ViewLoginCheck();
            return View();
        }
        public ActionResult QcNoticeItemSpec()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QcNoticeTemplate()
        {
            ViewLoginCheck();
            return View();
        }

        public ActionResult QcNoticeItemSpecTemplate()
        {
            ViewLoginCheck();
            return View();
        }


        #region//Get
        #region//GetQcGroup --取量測群組
        [HttpPost]
        public void GetQcGroup(int QcGroupId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read,constrained-data");

                #region //Request 
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcGroup(QcGroupId);
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

        #region//GetQcClass --取量測類別
        [HttpPost]
        public void GetQcClass(int QcGroupId = -1, int QcClassId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read,constrained-data");

                #region //Request 
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcClass(QcGroupId, QcClassId);
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

        #region//GetQcItem --取量測項目
        [HttpPost]
        public void GetQcItem(int QcItemId = -1, int QcClassId = -1, int QcGroupId = -1, string QcItemName = "", string SearchKey = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read,constrained-data");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcItem(QcItemId, QcClassId, QcGroupId, QcItemName, SearchKey
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

        #region//GetUnitOfQcMeasure --取得量測單位
        [HttpPost]
        public void GetUnitOfQcMeasure(int QcUomId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read,constrained-data");

                #region //Request 
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetUnitOfQcMeasure(QcUomId);
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

        #region//GetQcNotice --取得送檢單資料
        [HttpPost]
        public void GetQcNotice(int QcNoticeId=-1,string QcNoticeNo = "", string QcNoticeType = "",string WoErpPrefix = "",string WoErpNo = "",int WoSeq=-1,
            string MtlItemNo = "", string MtlItemName = "",int ProcessId=-1,
            string StartDate="",string EndDate="",int DepartmentId=-1,
            string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read,constrained-data");

                #region //Request   
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcNotice(QcNoticeId,QcNoticeNo, QcNoticeType, WoErpPrefix, WoErpNo, WoSeq, MtlItemNo, MtlItemName, ProcessId
                    , StartDate, EndDate, DepartmentId
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

        #region//GetQcNoticeItemSpec --取得送檢單量測項目資料 --Ding 2022-12-06
        [HttpPost]
        public void GetQcNoticeItemSpec(int QcNoticeItemSpecId = -1,int QcNoticeId=-1,
            string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read,constrained-data");

                #region //Request 
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcNoticeItemSpec(QcNoticeItemSpecId, QcNoticeId
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

        #region//GetQcNoticeFile --取得製程量測需求單-上傳檔案資料 --Ding 2022-12-08
        [HttpPost]
        public void GetQcNoticeFile(int QcNoticeId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read");

                #region //Request 
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcNoticeFile(QcNoticeId);
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

        #region//GetQcNoticeTemplate --取得送檢單樣板資料
        [HttpPost]
        public void GetQcNoticeTemplate(int QcNoticeTemplateId = -1, string QcNoticeTemplateNo = "", string QcNoticeType = "", int ParameterId = -1,
                    string StartDate = "", string EndDate = "",
                    string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "read");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcNoticeTemplate(QcNoticeTemplateId, QcNoticeTemplateNo, QcNoticeType, ParameterId
                    , StartDate, EndDate
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

        #region//GetQcNoticeItemSpecTemplate --取得樣板量測項目清單 --Ding 2022-12-07
        [HttpPost]
        public void GetQcNoticeItemSpecTemplate(int QcNoticeTemplateId = -1,int QcNoticeItemSpecTemplateId=-1,
            string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "read");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetQcNoticeItemSpecTemplate(QcNoticeTemplateId, QcNoticeItemSpecTemplateId
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

        #region//GetRoutingProcess --取得依製令對照途程製程
        [HttpPost]
        public void GetRoutingProcess(int ControlId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read");

                #region //Request 
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetRoutingProcess(ControlId);
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

        #region//GetRdInformation --取研發資訊
        [HttpPost]
        public void GetRdInformation(int ControlId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "read");

                #region //Request  
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetRdInformation(ControlId);
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

        #region //GetRdDesignControl 取得研發設計圖版本控制資料
        [HttpPost]
        public void GetRdDesignControl(int ControlId = -1, int DesignId = -1, string Edition = "", string StartDate = "", string EndDate = "", int MtlItemId = -1, string ReleasedStatus = ""
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                WebApiLoginCheck("RdDesign", "version");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetRdDesignControl(ControlId, DesignId, Edition, StartDate, EndDate, MtlItemId, ReleasedStatus
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

        #region//GetManufactureOrder 取得製令資訊
        [HttpPost]
        public void GetManufactureOrder(int MoId = -1, string otherInfo="")
        {
            try
            {
                //WebApiLoginCheck("MfgOrderManagement", "read,constrained-data");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.GetManufactureOrder(MoId, otherInfo);
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

        #region//Add

        #region//AddQcNotice --量測需求單單頭 新增 -- Ding 2022.12.06
        [HttpPost]
        public void AddQcNotice(string QcNoticeNo="", string QcNoticeType = "", int QcNoticeQty = -1, int MoId = -1,int ControlId=-1,int RoutingProcessId = -1,
            string Remark = "", string Status = "",string FileList="", int DepartmentId = -1, string ResetFlag = "")
        {
            try
            {
                WebApiLoginCheck("QcNotice", "add");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.AddQcNotice(QcNoticeNo,QcNoticeType,QcNoticeQty,MoId,ControlId,RoutingProcessId,Remark,Status, FileList, DepartmentId, ResetFlag);
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

        #region//AddQcNoticeItemSpec --製程量測需求單-量測項目清單 新增 -Ding 2022.12.06
        [HttpPost]
        public void AddQcNoticeItemSpec(int QcNoticeId = -1, int QcItemId = -1, double DesignValue = -1, double UpperTolerance = -1, double LowerTolerance = -1, int MakeCount = -1, int ProductQcCount = -1,
            int QcUomId=-1,int Depth=-1,string Remark="")
        {
            try
            {
                WebApiLoginCheck("QcNotice", "add");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.AddQcNoticeItemSpec(QcNoticeId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, MakeCount, ProductQcCount,
                    QcUomId, Depth, Remark);
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

        #region//AddQcNoticeTemplate --量測需求單樣板單頭 新增 -- Ding 2022.12.15
        [HttpPost]
        public void AddQcNoticeTemplate(string QcNoticeTemplateNo="", string QcNoticeTemplateName = "", string QcNoticeTemplateDesc = "", string QcNoticeType = "", int ParameterId = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "add");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.AddQcNoticeTemplate( QcNoticeTemplateNo,  QcNoticeTemplateName,  QcNoticeTemplateDesc,  QcNoticeType,  ParameterId,  Remark);
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

        #region//AddQcNoticeItemSpecTemplate --製程量測需求單樣板-量測項目清單 新增 -Ding 2022.12.06
        [HttpPost]
        public void AddQcNoticeItemSpecTemplate(int QcNoticeTemplateId = -1, int QcItemId = -1, double DesignValue = -1, double UpperTolerance = -1, double LowerTolerance = -1, int MakeCount = -1, int ProductQcCount = -1,
            int QcUomId = -1, int Depth = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "add");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.AddQcNoticeItemSpecTemplate(QcNoticeTemplateId, QcItemId, DesignValue, UpperTolerance, LowerTolerance, MakeCount, ProductQcCount,
                    QcUomId, Depth, Remark);
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

        #region//AddTemplateToQcNoticeItemSpec --樣板匯入
        [HttpPost]
        public void AddTemplateToQcNoticeItemSpec(int QcNoticeId = -1, int QcNoticeTemplateId = -1)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "add");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.AddTemplateToQcNoticeItemSpec(QcNoticeId, QcNoticeTemplateId);
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

        #region//Update

        #region//UpdateQcNotice --量測需求單單頭 修改 -- Ding 2022.12.07
        [HttpPost]
        public void UpdateQcNotice(int QcNoticeId = -1, string QcNoticeNo = "", string QcNoticeType = "", int QcNoticeQty = -1, int MoId = -1, int ControlId = -1, int RoutingProcessId = -1,
            int MoProcessId = -1, string Remark = "", string Status = "",string FileList="", int DepartmentId = -1, string ResetFlag = "")
        {
            try
            {
                WebApiLoginCheck("QcNotice", "update");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.UpdateQcNotice(QcNoticeId,QcNoticeNo,QcNoticeType,QcNoticeQty,MoId,ControlId,RoutingProcessId,
                    MoProcessId,Remark,Status,FileList,DepartmentId,ResetFlag);
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

        #region//UpdateQcNoticeItemSpec --製程量測需求單-量測項目清單 修改 --Ding 2022.10.17
        [HttpPost]
        public void UpdateQcNoticeItemSpec(int QcNoticeItemSpecId = -1, int QcNoticeId=-1, int QcItemId = -1, double DesignValue = -1, double UpperTolerance = -1, double LowerTolerance = -1
            , int MakeCount = -1,int ProductQcCount=-1, int QcUomId = -1, int Depth = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcNotice", "update");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.UpdateQcNoticeItemSpec(QcNoticeItemSpecId, QcNoticeId, QcItemId, DesignValue, UpperTolerance, LowerTolerance
                    , MakeCount, ProductQcCount, QcUomId, Depth, Remark);
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

        #region//UpdateQcNoticeTemplate --量測需求單單頭 修改 -- Ding 2022.12.07
        [HttpPost]
        public void UpdateQcNoticeTemplate(int QcNoticeTemplateId = -1, string QcNoticeTemplateNo = "", string QcNoticeTemplateName = "", string QcNoticeTemplateDesc = "", string QcNoticeType = "", int ParameterId = -1,
            string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "update");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.UpdateQcNoticeTemplate(QcNoticeTemplateId, QcNoticeTemplateNo, QcNoticeTemplateName, QcNoticeTemplateDesc, QcNoticeType, ParameterId,
                   Remark);
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

        #region//UpdateQcNoticeItemSpecTemplate --製程量測需求單-量測項目清單 修改 --Ding 2022.10.17
        [HttpPost]
        public void UpdateQcNoticeItemSpecTemplate(int QcNoticeItemSpecTemplateId = -1, int QcNoticeTemplateId = -1, int QcItemId = -1, double DesignValue = -1, double UpperTolerance = -1, double LowerTolerance = -1
            , int MakeCount = -1, int ProductQcCount = -1, int QcUomId = -1, int Depth = -1, string Remark = "")
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "update");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.UpdateQcNoticeItemSpecTemplate(QcNoticeItemSpecTemplateId, QcNoticeTemplateId, QcItemId, DesignValue, UpperTolerance, LowerTolerance
                    , MakeCount, ProductQcCount, QcUomId, Depth, Remark);
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

        #region//Delete
        #region//DeleteQcNotice --送檢單資料 刪除 --Ding 2022.10.17
        [HttpPost]
        public void DeleteQcNotice(int QcNoticeId)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "delete");

                #region //Request  
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.DeleteQcNotice(QcNoticeId);
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

        #region//DeleteQcNoticeItemSpec --製程量測需求單-量測項目清單 刪除 --Ding 2022.10.17
        [HttpPost]
        public void DeleteQcNoticeItemSpec(int QcNoticeItemSpecId)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "delete");

                #region //Request  
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.DeleteQcNoticeItemSpec(QcNoticeItemSpecId);
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

        #region//DeleteQcNoticeFile --送檢單附件 刪除 --Ding 2022.12.8 
        [HttpPost]
        public void DeleteQcNoticeFile(int QcNoticeFileId)
        {
            try
            {
                WebApiLoginCheck("QcNotice", "delete");

                #region //Request
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.DeleteQcNoticeFile(QcNoticeFileId);
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

        #region//DeleteQcNoticeTemplate --送檢單資料樣板 刪除 --Ding 2022.10.17
        [HttpPost]
        public void DeleteQcNoticeTemplate(int QcNoticeTemplateId)
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "delete");

                #region //Request  
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.DeleteQcNoticeTemplate(QcNoticeTemplateId);
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

        #region//DeleteQcNoticeItemSpecTemplate --製程量測需求單-量測項目清單-樣板 刪除 --Ding 2022.10.17
        [HttpPost]
        public void DeleteQcNoticeItemSpecTemplate(int QcNoticeItemSpecTemplateId)
        {
            try
            {
                WebApiLoginCheck("QcNoticeTemplate", "delete");

                #region //Request   
                qcItemSettingDA = new QcItemSettingDA();
                dataRequest = qcItemSettingDA.DeleteQcNoticeItemSpecTemplate(QcNoticeItemSpecTemplateId);
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